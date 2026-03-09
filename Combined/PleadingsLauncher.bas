Attribute VB_Name = "PleadingsLauncher"
' ============================================================
' PleadingsLauncher.bas
' Lightweight launcher for the Pleadings Checker.
' Uses MsgBox/InputBox only -- no UserForm required.
'
' Dependencies:
'   - PleadingsEngine.bas
'   - PleadingsIssue.cls
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
    Dim cfg As Scripting.Dictionary
    Set cfg = PleadingsEngine.InitRuleConfig()

    ' -- Page range prompt --
    Dim pgInput As String
    pgInput = InputBox("Page range (e.g. 1-10, or leave blank for all pages):", _
                        "Pleadings Checker - Page Range", "")
    If Len(Trim(pgInput)) > 0 Then
        ParsePageRange pgInput
    Else
        PleadingsEngine.SetPageRange 0, 0
    End If

    ' -- Spelling mode prompt --
    Dim spMode As Long
    spMode = MsgBox("Enforce UK spelling?" & vbCrLf & vbCrLf & _
                    "Yes = UK spelling (default)" & vbCrLf & _
                    "No = US spelling", _
                    vbYesNo + vbQuestion, "Spelling Mode")
    If spMode = vbNo Then
        PleadingsEngine.SetSpellingMode "US"
    Else
        PleadingsEngine.SetSpellingMode "UK"
    End If

    ' -- Run --
    Application.StatusBar = "Pleadings Checker: running checks..."
    DoEvents

    Dim issues As Collection
    Set issues = PleadingsEngine.RunAllPleadingsRules(ActiveDocument, cfg)

    Application.StatusBar = ""

    ' -- Show results --
    If issues.Count = 0 Then
        MsgBox "No issues found " & Chr(8212) & " document looks clean.", _
               vbInformation, "Pleadings Checker"
        Exit Sub
    End If

    Dim summary As String
    summary = PleadingsEngine.GetIssueSummary(issues)

    Dim applyChoice As Long
    applyChoice = MsgBox(summary & vbCrLf & vbCrLf & _
                         "Apply to document?" & vbCrLf & _
                         "Yes = Apply as tracked changes" & vbCrLf & _
                         "No = Highlight + comments only" & vbCrLf & _
                         "Cancel = View results only", _
                         vbYesNoCancel + vbInformation, _
                         "Pleadings Checker " & Chr(8212) & " " & _
                         issues.Count & " Issue(s)")

    Select Case applyChoice
        Case vbYes
            PleadingsEngine.ApplySuggestionsAsTrackedChanges ActiveDocument, issues, True
            MsgBox issues.Count & " issue(s) applied as tracked changes.", _
                   vbInformation, "Pleadings Checker"
        Case vbNo
            PleadingsEngine.ApplyHighlights ActiveDocument, issues, True
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
            loadPath = Environ("APPDATA") & "\PleadingsChecker\brand_rules.txt"
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
            savePath = Environ("APPDATA") & "\PleadingsChecker\brand_rules.txt"
            On Error Resume Next
            MkDir Environ("APPDATA") & "\PleadingsChecker"
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
    If ActiveDocument.Path <> "" Then
        reportPath = ActiveDocument.Path & "\" & _
                     Replace(ActiveDocument.Name, ".docx", "") & _
                     "_pleadings_report.json"
    Else
        reportPath = Environ("TEMP") & "\pleadings_report.json"
    End If

    Dim summary As String
    summary = PleadingsEngine.GenerateReport(issues, reportPath)

    MsgBox "Report saved to:" & vbCrLf & reportPath, _
           vbInformation, "Pleadings Checker"
End Sub

' ============================================================
'  PARSE PAGE RANGE INPUT (e.g. "1-10" or "5")
' ============================================================
Private Sub ParsePageRange(ByVal input As String)
    Dim parts() As String
    Dim startPg As Long
    Dim endPg As Long

    input = Trim(input)
    If InStr(1, input, "-") > 0 Then
        parts = Split(input, "-")
        If UBound(parts) >= 1 Then
            If IsNumeric(Trim(parts(0))) And IsNumeric(Trim(parts(1))) Then
                startPg = CLng(Trim(parts(0)))
                endPg = CLng(Trim(parts(1)))
                PleadingsEngine.SetPageRange startPg, endPg
                Exit Sub
            End If
        End If
    ElseIf IsNumeric(input) Then
        startPg = CLng(input)
        PleadingsEngine.SetPageRange startPg, startPg
        Exit Sub
    End If

    ' Invalid input -- use all pages
    PleadingsEngine.SetPageRange 0, 0
End Sub
