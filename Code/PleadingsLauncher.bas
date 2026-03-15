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

Private targetDoc As Document

' ============================================================
'  MAIN LAUNCHER (called by PleadingsEngine.PleadingsChecker)
'  User-facing entry point: prompts for target document, then
'  offers Run / Options / Cancel.  All subsequent operations
'  use the selected target document, not ActiveDocument.
' ============================================================
Public Sub LaunchChecker()
    Set targetDoc = Application.Run("PleadingsEngine.GetTargetDocument")
    If targetDoc Is Nothing Then
        Exit Sub
    End If

    ' -- Choose action --
    Dim choice As Long
    choice = MsgBox("Pleadings Checker" & vbCrLf & vbCrLf & _
                    "Document: " & targetDoc.Name & vbCrLf & vbCrLf & _
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
    Application.Run "PleadingsEngine.SetPageRangeFromString", Trim(pgInput)

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
    Set issues = Application.Run("PleadingsEngine.RunAllPleadingsRules", targetDoc, cfg)

    Application.StatusBar = ""

    ' -- Show results --
    Dim errCount As Long
    errCount = Application.Run("PleadingsEngine.GetRuleErrorCount")

    If issues.Count = 0 Then
        If errCount > 0 Then
            Dim errLog As String
            errLog = Application.Run("PleadingsEngine.GetRuleErrorLog")
            MsgBox "No issues found, but " & errCount & " rule(s) failed to run:" & vbCrLf & vbCrLf & _
                   errLog & vbCrLf & _
                   "Check Immediate window (Ctrl+G) or export a report for the debug log.", _
                   vbExclamation, "Pleadings Checker"
        Else
            MsgBox "No issues found -- document looks clean.", _
                   vbInformation, "Pleadings Checker"
        End If
        Exit Sub
    End If

    Dim errInfo As String
    If errCount > 0 Then
        errInfo = vbCrLf & errCount & " rule(s) failed to run."
    End If

    Dim issueSummary As String
    issueSummary = Application.Run("PleadingsEngine.GetIssueSummary", issues)

    Dim applyChoice As Long
    applyChoice = MsgBox(issueSummary & errInfo & vbCrLf & vbCrLf & _
                         "Apply to document?" & vbCrLf & _
                         "Yes = Apply as tracked changes" & vbCrLf & _
                         "No = Highlight + comments only" & vbCrLf & _
                         "Cancel = View results only", _
                         vbYesNoCancel + vbInformation, _
                         "Pleadings Checker -- " & _
                         issues.Count & " Issue(s)")

    Select Case applyChoice
        Case vbYes
            Application.Run "PleadingsEngine.ApplySuggestionsAsTrackedChanges", targetDoc, issues, True
            MsgBox issues.Count & " issue(s) applied as tracked changes.", _
                   vbInformation, "Pleadings Checker"
        Case vbNo
            Application.Run "PleadingsEngine.ApplyHighlights", targetDoc, issues, True
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
            loadPath = GetBrandRulesPath()
            On Error Resume Next
            Dim loadOK As Boolean
            loadOK = Application.Run("Rules_Brands.LoadBrandRules", loadPath)
            If Err.Number <> 0 Then
                MsgBox "Rules_Brands module not imported or file not found." & vbCrLf & _
                       "Error: " & Err.Description, vbExclamation, "Pleadings Checker"
                Err.Clear
            ElseIf loadOK Then
                MsgBox "Brand rules loaded.", vbInformation, "Pleadings Checker"
            Else
                MsgBox "Brand rules file could not be read:" & vbCrLf & loadPath, _
                       vbExclamation, "Pleadings Checker"
            End If
            On Error GoTo 0

        Case "SAVE"
            Dim savePath As String
            savePath = GetBrandRulesPath()
            ' Ensure directory exists (recursive, handles nested paths)
            Dim brandDir As String
            brandDir = GetParentDirectory(savePath)
            If Len(brandDir) > 0 Then
                EnsureDirectoryExists brandDir
            End If
            On Error Resume Next
            Dim saveOK As Boolean
            saveOK = Application.Run("Rules_Brands.SaveBrandRules", savePath)
            If Err.Number <> 0 Then
                MsgBox "Rules_Brands module not imported." & vbCrLf & _
                       "Error: " & Err.Description, vbExclamation, "Pleadings Checker"
                Err.Clear
            ElseIf saveOK Then
                MsgBox "Brand rules saved to:" & vbCrLf & savePath, _
                       vbInformation, "Pleadings Checker"
            Else
                MsgBox "Failed to save brand rules to:" & vbCrLf & savePath, _
                       vbExclamation, "Pleadings Checker"
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

    On Error Resume Next
    If Not targetDoc Is Nothing Then
        If targetDoc.Path <> "" Then
            Dim baseName As String
            baseName = targetDoc.Name
            Dim dotPos As Long
            dotPos = InStrRev(baseName, ".")
            If dotPos > 1 Then baseName = Left$(baseName, dotPos - 1)
            reportPath = targetDoc.Path & sep & baseName & "_pleadings_report.json"
        End If
    End If
    If Err.Number <> 0 Or Len(reportPath) = 0 Then
        Err.Clear
        reportPath = ""
    End If
    On Error GoTo 0

    If Len(reportPath) = 0 Then
        reportPath = GetWritableTempDir() & sep & "pleadings_report.json"
    End If

    ' Ensure parent directory exists before writing
    Dim reportDir As String
    reportDir = GetParentDirectory(reportPath)
    If Len(reportDir) > 0 Then
        EnsureDirectoryExists reportDir
    End If

    Dim summary As String
    summary = Application.Run("PleadingsEngine.GenerateReport", issues, reportPath, targetDoc)

    ' Auto-save debug log alongside report when DEBUG_MODE is True
    Dim logPath As String
    Dim logSaved As Boolean
    logSaved = False
    logPath = ""

    On Error Resume Next
    If DEBUG_MODE Then
        logPath = Left$(reportPath, Len(reportPath) - 5) & "_debug.log"
        logSaved = DebugLogSaveToTextFile(logPath)
    End If
    On Error GoTo 0

    Dim msg As String
    msg = "Report saved to:" & vbCrLf & reportPath

    If logSaved And Len(logPath) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Debug log saved to:" & vbCrLf & logPath
    ElseIf DEBUG_MODE And Not logSaved Then
        msg = msg & vbCrLf & vbCrLf & "Debug log could not be saved."
    End If

    Dim exportErrCount As Long
    On Error Resume Next
    exportErrCount = Application.Run("PleadingsEngine.GetRuleErrorCount")
    If Err.Number <> 0 Then exportErrCount = 0: Err.Clear
    On Error GoTo 0
    If exportErrCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & exportErrCount & " rule(s) failed during the run."
    End If

    MsgBox msg, vbInformation, "Pleadings Checker"
End Sub

' ============================================================
'  PRIVATE: Cross-platform brand rules file path
'  Delegates to Rules_Brands.GetDefaultBrandRulesPath (single source of truth).
'  Falls back to a local construction if the module is not imported.
' ============================================================
Private Function GetBrandRulesPath() As String
    On Error Resume Next
    GetBrandRulesPath = Application.Run("Rules_Brands.GetDefaultBrandRulesPath")
    If Err.Number <> 0 Then
        Debug.Print "GetBrandRulesPath: Rules_Brands not loaded (Err " & Err.Number & "); using inline fallback"
        Err.Clear
        On Error GoTo 0
        ' Fallback: build the path locally (kept in sync with Rules_Brands.GetDefaultBrandRulesPath)
        Dim sep As String
        sep = Application.PathSeparator
        #If Mac Then
            GetBrandRulesPath = Environ("HOME") & sep & "Library" & sep & _
                                "Application Support" & sep & "PleadingsChecker" & sep & "brand_rules.txt"
        #Else
            GetBrandRulesPath = Environ("APPDATA") & sep & "PleadingsChecker" & sep & "brand_rules.txt"
        #End If
        Exit Function
    End If
    On Error GoTo 0
End Function

