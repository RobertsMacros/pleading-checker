VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPleadingsChecker
   Caption         =   "Pleadings Checker"
   ClientHeight    =   13200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8400
   OleObjectBlob   =   "frmPleadingsChecker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPleadingsChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ============================================================
' frmPleadingsChecker.frm
' UserForm for the Pleadings Checker rule engine.
' Provides: dynamic scrollable rule checkbox list, brand rule
' management, page range config, tracked-changes suggestion
' mode, and run/report controls.
'
' The rule list is generated dynamically from the engine's
' rule metadata so that adding new rules requires no form
' changes — only engine + module additions.
' ============================================================
Option Explicit

' ── Module-level variables ────────────────────────────────────
Private ruleConfig      As Object  ' Scripting.Dictionary
Private ruleDisplayMap  As Object  ' Scripting.Dictionary (rule_name -> label)
Private ruleKeys()      As String  ' Ordered array of rule names
Private ruleCheckboxes  As Collection  ' MSForms.CheckBox controls
Private lastResults     As Collection

' ════════════════════════════════════════════════════════════
'  FORM INITIALISATION
' ════════════════════════════════════════════════════════════
Private Sub UserForm_Initialize()
    ' Build rule config and display names from engine
    Set ruleConfig = PleadingsEngine.InitRuleConfig()
    Set ruleDisplayMap = PleadingsEngine.GetRuleDisplayNames()

    ' Build ordered key array from config (preserves insertion order)
    Dim keys As Variant
    keys = ruleConfig.keys
    Dim nRules As Long
    nRules = ruleConfig.Count
    ReDim ruleKeys(0 To nRules - 1)
    Dim k As Long
    For k = 0 To nRules - 1
        ruleKeys(k) = CStr(keys(k))
    Next k

    ' ── Build scrollable rule checkbox list ─────────────────────
    BuildRuleCheckboxList nRules

    ' ── Page range defaults ───────────────────────────────────
    txtStartPage.Text = ""
    txtEndPage.Text = ""

    ' ── Brand rules ───────────────────────────────────────────
    RefreshBrandList

    ' ── Tracked changes checkbox default ──────────────────────
    chkTrackedChanges.value = True

    ' ── Status ────────────────────────────────────────────────
    lblStatus.Caption = "Ready. Select rules and click Run."
End Sub

' ════════════════════════════════════════════════════════════
'  BUILD DYNAMIC RULE CHECKBOX LIST
'  Creates checkboxes inside fraRules (a scrollable frame)
'  one per rule, with labels derived from GetRuleDisplayNames.
' ════════════════════════════════════════════════════════════
Private Sub BuildRuleCheckboxList(nRules As Long)
    Set ruleCheckboxes = New Collection

    Dim topPos As Single
    Dim chk As MSForms.CheckBox
    Dim displayLabel As String
    Dim i As Long

    topPos = 6  ' initial top padding inside frame

    For i = 0 To nRules - 1
        ' Determine display label
        If ruleDisplayMap.Exists(ruleKeys(i)) Then
            displayLabel = CStr(i + 1) & ". " & CStr(ruleDisplayMap(ruleKeys(i)))
        Else
            displayLabel = CStr(i + 1) & ". " & ruleKeys(i)
        End If

        ' Create checkbox control inside the scrollable frame
        Set chk = fraRules.Controls.Add("Forms.CheckBox.1", "chkRule_" & i)
        With chk
            .Caption = displayLabel
            .Left = 6
            .Top = topPos
            .Width = fraRules.InsideWidth - 18
            .Height = 18
            .value = True
            .Font.Size = 9
        End With

        ruleCheckboxes.Add chk
        topPos = topPos + 20
    Next i

    ' Set scroll height to fit all checkboxes
    fraRules.ScrollHeight = topPos + 6
    fraRules.ScrollBars = fmScrollBarsVertical
End Sub

' ════════════════════════════════════════════════════════════
'  RUN BUTTON
' ════════════════════════════════════════════════════════════
Private Sub btnRun_Click()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Pleadings Checker"
        Exit Sub
    End If

    ' ── Sync rule config from dynamic checkboxes ────────────────
    Dim i As Long
    For i = 1 To ruleCheckboxes.Count
        Dim rName As String
        rName = ruleKeys(i - 1)
        If ruleConfig.Exists(rName) Then
            ruleConfig(rName) = CBool(ruleCheckboxes(i).value)
        End If
    Next i

    ' ── Set page range if specified ───────────────────────────
    Dim startPg As Long, endPg As Long
    startPg = 0: endPg = 0
    If IsNumeric(txtStartPage.Text) And Len(txtStartPage.Text) > 0 Then
        startPg = CLng(txtStartPage.Text)
    End If
    If IsNumeric(txtEndPage.Text) And Len(txtEndPage.Text) > 0 Then
        endPg = CLng(txtEndPage.Text)
    End If
    PleadingsEngine.SetPageRange startPg, endPg

    ' ── Run checks ────────────────────────────────────────────
    lblStatus.Caption = "Running checks..."
    Me.Repaint
    DoEvents

    Set lastResults = PleadingsEngine.RunAllPleadingsRules(ActiveDocument, ruleConfig)

    ' ── Show summary ──────────────────────────────────────────
    Dim summary As String
    summary = PleadingsEngine.GetIssueSummary(lastResults)

    If lastResults.Count = 0 Then
        lblStatus.Caption = "No issues found."
        MsgBox "No issues found " & Chr(8212) & " document looks clean.", vbInformation, "Pleadings Checker"
    Else
        lblStatus.Caption = lastResults.Count & " issue(s) found."
        MsgBox summary, vbInformation, "Pleadings Checker " & Chr(8212) & " Results"
    End If
End Sub

' ════════════════════════════════════════════════════════════
'  HIGHLIGHT / APPLY SUGGESTIONS BUTTON
'  If tracked-changes mode is on, uses the tracked-changes
'  applier so auto-fix-safe suggestions become revisions.
'  Otherwise uses legacy highlight+comment mode.
' ════════════════════════════════════════════════════════════
Private Sub btnHighlight_Click()
    If lastResults Is Nothing Then
        MsgBox "Run checks first before highlighting.", vbExclamation, "Pleadings Checker"
        Exit Sub
    End If
    If lastResults.Count = 0 Then
        MsgBox "No issues to highlight.", vbInformation, "Pleadings Checker"
        Exit Sub
    End If

    Dim addComments As Boolean
    addComments = (chkAddComments.value = True)

    lblStatus.Caption = "Applying suggestions..."
    Me.Repaint
    DoEvents

    If chkTrackedChanges.value = True Then
        ' Use tracked-changes mode: auto-fix-safe -> revisions, others -> comments
        PleadingsEngine.ApplySuggestionsAsTrackedChanges ActiveDocument, lastResults, addComments
    Else
        ' Legacy mode: highlight + optional comments only
        PleadingsEngine.ApplyHighlights ActiveDocument, lastResults, addComments
    End If

    lblStatus.Caption = lastResults.Count & " issue(s) processed."
    MsgBox lastResults.Count & " issue(s) processed in document." & vbCrLf & _
           IIf(chkTrackedChanges.value, "Auto-fix suggestions applied as tracked changes.", _
               "Issues highlighted.") & vbCrLf & _
           IIf(addComments, "Comments added for non-auto-fix items.", ""), _
           vbInformation, "Pleadings Checker"
End Sub

' ════════════════════════════════════════════════════════════
'  EXPORT REPORT BUTTON
' ════════════════════════════════════════════════════════════
Private Sub btnExport_Click()
    If lastResults Is Nothing Then
        MsgBox "Run checks first before exporting.", vbExclamation, "Pleadings Checker"
        Exit Sub
    End If

    ' Build report path next to the document
    Dim reportPath As String
    If ActiveDocument.Path <> "" Then
        reportPath = ActiveDocument.Path & "\" & _
                     Replace(ActiveDocument.Name, ".docx", "") & _
                     "_pleadings_report.json"
    Else
        reportPath = Environ("TEMP") & "\pleadings_report.json"
    End If

    lblStatus.Caption = "Exporting report..."
    Me.Repaint
    DoEvents

    Dim summary As String
    summary = PleadingsEngine.GenerateReport(lastResults, reportPath)

    lblStatus.Caption = "Report saved."
    MsgBox "Report saved to:" & vbCrLf & reportPath & vbCrLf & vbCrLf & summary, _
           vbInformation, "Pleadings Checker " & Chr(8212) & " Report"
End Sub

' ════════════════════════════════════════════════════════════
'  SELECT ALL / DESELECT ALL
' ════════════════════════════════════════════════════════════
Private Sub btnSelectAll_Click()
    Dim i As Long
    For i = 1 To ruleCheckboxes.Count
        ruleCheckboxes(i).value = True
    Next i
End Sub

Private Sub btnDeselectAll_Click()
    Dim i As Long
    For i = 1 To ruleCheckboxes.Count
        ruleCheckboxes(i).value = False
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  BRAND RULES MANAGEMENT
' ════════════════════════════════════════════════════════════
Private Sub RefreshBrandList()
    lstBrands.Clear
    Dim brands As Object
    Set brands = Rule22_brand_name_enforcement.GetBrandRules()
    If brands Is Nothing Then Exit Sub
    Dim key As Variant
    For Each key In brands.keys
        lstBrands.AddItem CStr(key) & " -> " & CStr(brands(key))
    Next key
End Sub

Private Sub btnAddBrand_Click()
    Dim correctForm As String
    Dim incorrectVariants As String
    correctForm = Trim(txtBrandCorrect.Text)
    incorrectVariants = Trim(txtBrandIncorrect.Text)

    If correctForm = "" Or incorrectVariants = "" Then
        MsgBox "Enter both the correct form and incorrect variants.", _
               vbExclamation, "Brand Rules"
        Exit Sub
    End If

    Rule22_brand_name_enforcement.AddBrandRule correctForm, incorrectVariants
    txtBrandCorrect.Text = ""
    txtBrandIncorrect.Text = ""
    RefreshBrandList
End Sub

Private Sub btnRemoveBrand_Click()
    If lstBrands.ListIndex < 0 Then
        MsgBox "Select a brand rule to remove.", vbExclamation, "Brand Rules"
        Exit Sub
    End If

    Dim entry As String
    entry = lstBrands.List(lstBrands.ListIndex)
    ' Extract correct form (before " -> ")
    Dim correctForm As String
    correctForm = Left(entry, InStr(entry, " -> ") - 1)

    Rule22_brand_name_enforcement.RemoveBrandRule correctForm
    RefreshBrandList
End Sub

Private Sub btnSaveBrands_Click()
    Dim brandFile As String
    brandFile = Environ("APPDATA") & "\PleadingsChecker\brand_rules.txt"

    ' Create directory if needed
    On Error Resume Next
    MkDir Environ("APPDATA") & "\PleadingsChecker"
    On Error GoTo 0

    Rule22_brand_name_enforcement.SaveBrandRules brandFile
    MsgBox "Brand rules saved to:" & vbCrLf & brandFile, vbInformation, "Brand Rules"
End Sub

Private Sub btnLoadBrands_Click()
    Dim brandFile As String
    brandFile = Environ("APPDATA") & "\PleadingsChecker\brand_rules.txt"

    If Dir(brandFile) = "" Then
        MsgBox "No saved brand rules found at:" & vbCrLf & brandFile, _
               vbExclamation, "Brand Rules"
        Exit Sub
    End If

    Rule22_brand_name_enforcement.LoadBrandRules brandFile
    RefreshBrandList
    MsgBox "Brand rules loaded.", vbInformation, "Brand Rules"
End Sub

' ════════════════════════════════════════════════════════════
'  CLOSE BUTTON
' ════════════════════════════════════════════════════════════
Private Sub btnClose_Click()
    Unload Me
End Sub

' ════════════════════════════════════════════════════════════
'  FORM LAYOUT NOTES
' ════════════════════════════════════════════════════════════
' This form requires the following controls to be created manually
' in the VBA UserForm designer (or via .frx binary data):
'
' RULE SELECTION (top section — scrollable frame):
'   fraRules         - Frame, Caption="Rules", ScrollBars=fmScrollBarsVertical
'                      Position: Left=12, Top=12, Width=396, Height=300
'                      ScrollHeight set dynamically based on rule count
'   btnSelectAll     - CommandButton, Caption="Select All"
'                      Position: Left=420, Top=24, Width=96, Height=24
'   btnDeselectAll   - CommandButton, Caption="Deselect All"
'                      Position: Left=420, Top=54, Width=96, Height=24
'
'   Checkboxes are created dynamically inside fraRules by
'   BuildRuleCheckboxList — no manual checkbox creation needed.
'
' PAGE RANGE (below rules):
'   Label            - Caption="Start Page:"
'   txtStartPage     - TextBox, Width=48
'   Label            - Caption="End Page:"
'   txtEndPage       - TextBox, Width=48
'
' BRAND RULES (middle section):
'   lstBrands        - ListBox, Width=360, Height=96
'   Label            - Caption="Correct Form:"
'   txtBrandCorrect  - TextBox, Width=144
'   Label            - Caption="Incorrect Variants:"
'   txtBrandIncorrect - TextBox, Width=144
'   btnAddBrand      - CommandButton, Caption="Add"
'   btnRemoveBrand   - CommandButton, Caption="Remove"
'   btnSaveBrands    - CommandButton, Caption="Save Rules"
'   btnLoadBrands    - CommandButton, Caption="Load Rules"
'
' OPTIONS:
'   chkAddComments   - CheckBox, Caption="Add comments to document", Value=True
'   chkTrackedChanges - CheckBox, Caption="Apply suggestions as tracked changes", Value=True
'
' ACTION BUTTONS (bottom):
'   btnRun           - CommandButton, Caption="Run Checks", Width=96, Height=30
'   btnHighlight     - CommandButton, Caption="Apply Suggestions", Width=108, Height=30
'   btnExport        - CommandButton, Caption="Export Report", Width=96, Height=30
'   btnClose         - CommandButton, Caption="Close", Width=72, Height=30
'
' STATUS BAR:
'   lblStatus        - Label, Caption="Ready.", Width=504
