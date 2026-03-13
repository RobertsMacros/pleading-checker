# FILE: frmPleadingsChecker.frm

```vb
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPleadingsChecker
   Caption         =   "Pleadings Checker"
   ClientHeight    =   1000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1000
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
'
' ALL controls are created dynamically in UserForm_Initialize
' so that no .frx binary file is needed.  The form adapts its
' rule checkbox list from the engine metadata automatically.
'
' All controls are created dynamically below.
' ============================================================
Option Explicit

' -- Module-level variables ------------------------------------
Private ruleConfig      As Object  ' Scripting.Dictionary
Private ruleDisplayMap  As Object  ' Scripting.Dictionary (rule_name -> label)
Private ruleKeys()      As String  ' Ordered array of rule names
Private ruleCheckboxes  As Collection  ' MSForms.CheckBox controls

' Controls created at runtime (module-level so event subs can reference them)
Private WithEvents btnRun           As MSForms.CommandButton
Private WithEvents btnExport        As MSForms.CommandButton
Private WithEvents btnClose         As MSForms.CommandButton
Private WithEvents btnSelectAll     As MSForms.CommandButton
Private WithEvents btnDeselectAll   As MSForms.CommandButton
Private WithEvents btnAddBrand      As MSForms.CommandButton
Private WithEvents btnRemoveBrand   As MSForms.CommandButton
Private WithEvents btnSaveBrands    As MSForms.CommandButton
Private WithEvents btnLoadBrands    As MSForms.CommandButton

Private fraRules        As MSForms.Frame
Private txtPageRange    As MSForms.TextBox
Private lstBrands       As MSForms.ListBox
Private txtBrandCorrect As MSForms.TextBox
Private txtBrandIncorrect As MSForms.TextBox
Private chkAddComments  As MSForms.CheckBox
Private chkTrackedChanges As MSForms.CheckBox
Private optSpellingUK   As MSForms.OptionButton
Private optSpellingUS   As MSForms.OptionButton
Private optQuoteSingle  As MSForms.OptionButton
Private optQuoteDouble  As MSForms.OptionButton
Private optSmart   As MSForms.OptionButton
Private optSmartStraight As MSForms.OptionButton
Private optDateUK       As MSForms.OptionButton
Private optDateUS       As MSForms.OptionButton
Private cboTermFormat   As MSForms.ComboBox
Private cboTermQuotes   As MSForms.ComboBox
Private cboSpaceStyle   As MSForms.ComboBox
Private lblStatus       As MSForms.Label

Private lastResults     As Collection

' ============================================================
'  FORM INITIALISATION -- creates all controls at runtime
' ============================================================
Private Sub UserForm_Initialize()
    Dim lbl As MSForms.Label
    Dim yPos As Single

    ' -- Overall form padding ----------------------------------
    Const PAD As Single = 12
    Const FULL_W As Single = 976     ' usable width (form 1000 - 2*PAD)
    Const BTN_W As Single = 108
    Const BTN_H As Single = 26
    Const TXT_H As Single = 22
    Const CHK_H As Single = 18
    Const LBL_H As Single = 16
    Const SEC_GAP As Single = 10     ' gap between sections
    Const ITEM_GAP As Single = 4     ' gap within sections

    ' -- Build rule data first (need count for layout) ---------
    Set ruleConfig = PleadingsEngine.InitRuleConfig()
    Set ruleDisplayMap = PleadingsEngine.GetRuleDisplayNames()

    Dim keys As Variant
    keys = ruleConfig.keys
    Dim nRules As Long
    nRules = ruleConfig.Count
    ReDim ruleKeys(0 To nRules - 1)
    Dim k As Long
    For k = 0 To nRules - 1
        ruleKeys(k) = CStr(keys(k))
    Next k

    yPos = PAD

    ' ==========================================================
    '  ROW 1: Rules header + Select All / Deselect All inline
    ' ==========================================================
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblRulesHeader")
    With lbl
        .Caption = "Rules"
        .Left = PAD: .Top = yPos: .Width = 60: .Height = LBL_H
        .Font.Size = 10: .Font.Bold = True
    End With

    Set btnSelectAll = Me.Controls.Add("Forms.CommandButton.1", "btnSelectAll")
    With btnSelectAll
        .Caption = "Select All"
        .Left = PAD + 66: .Top = yPos - 2: .Width = 78: .Height = 22
        .Font.Size = 8
    End With

    Set btnDeselectAll = Me.Controls.Add("Forms.CommandButton.1", "btnDeselectAll")
    With btnDeselectAll
        .Caption = "Deselect All"
        .Left = PAD + 66 + 82: .Top = yPos - 2: .Width = 78: .Height = 22
        .Font.Size = 8
    End With

    yPos = yPos + 22 + ITEM_GAP

    ' ==========================================================
    '  ROW 2: Rule checkboxes in multi-column scrollable frame
    ' ==========================================================
    Set fraRules = Me.Controls.Add("Forms.Frame.1", "fraRules")
    With fraRules
        .Caption = ""
        .Left = PAD: .Top = yPos
        .Width = FULL_W
        .Height = 120
        .ScrollBars = fmScrollBarsVertical
        .KeepScrollBarsVisible = fmScrollBarsVertical
    End With

    BuildRuleCheckboxList nRules

    yPos = yPos + fraRules.Height + SEC_GAP

    ' ==========================================================
    '  ROW 3: Page Range + Options side by side
    ' ==========================================================
    Dim colLeft As Single
    Dim colRight As Single
    colLeft = PAD
    colRight = PAD + FULL_W / 2 + SEC_GAP

    ' -- Left: Page Range --
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblPageHeader")
    With lbl
        .Caption = "Page Range (optional)"
        .Left = colLeft: .Top = yPos: .Width = 200: .Height = LBL_H
        .Font.Size = 10: .Font.Bold = True
    End With

    ' -- Right: Options --
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblOptionsHeader")
    With lbl
        .Caption = "Options"
        .Left = colRight: .Top = yPos: .Width = 200: .Height = LBL_H
        .Font.Size = 10: .Font.Bold = True
    End With
    yPos = yPos + LBL_H + ITEM_GAP

    ' Page range field (flexible format: "5", "3-7", "1,3,5", "1,3-5,8")
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblPageRange")
    With lbl
        .Caption = "Pages:"
        .Left = colLeft: .Top = yPos + 3: .Width = 40: .Height = LBL_H
    End With

    Set txtPageRange = Me.Controls.Add("Forms.TextBox.1", "txtPageRange")
    With txtPageRange
        .Left = colLeft + 40: .Top = yPos: .Width = 196: .Height = TXT_H
        .Text = ""
    End With

    ' Options checkboxes (right column, same rows)
    Set chkAddComments = Me.Controls.Add("Forms.CheckBox.1", "chkAddComments")
    With chkAddComments
        .Caption = "Add comments to document"
        .Left = colRight: .Top = yPos: .Width = 240: .Height = CHK_H
        .Value = True
    End With
    yPos = yPos + TXT_H + ITEM_GAP

    Set chkTrackedChanges = Me.Controls.Add("Forms.CheckBox.1", "chkTrackedChanges")
    With chkTrackedChanges
        .Caption = "Apply suggestions as tracked changes"
        .Left = colRight: .Top = yPos: .Width = 280: .Height = CHK_H
        .Value = True
    End With
    yPos = yPos + CHK_H + ITEM_GAP

    ' -- Spelling mode toggle (UK / US) --
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblSpellingMode")
    With lbl
        .Caption = "Spelling mode:"
        .Left = colRight: .Top = yPos + 2: .Width = 80: .Height = LBL_H
    End With

    Set optSpellingUK = Me.Controls.Add("Forms.OptionButton.1", "optSpellingUK")
    With optSpellingUK
        .Caption = "UK"
        .Left = colRight + 82: .Top = yPos: .Width = 50: .Height = CHK_H
        .Value = True
        .GroupName = "SpellingMode"
    End With

    Set optSpellingUS = Me.Controls.Add("Forms.OptionButton.1", "optSpellingUS")
    With optSpellingUS
        .Caption = "US"
        .Left = colRight + 134: .Top = yPos: .Width = 50: .Height = CHK_H
        .Value = False
        .GroupName = "SpellingMode"
    End With

    yPos = yPos + CHK_H + ITEM_GAP

    ' -- Quote nesting toggle (Single outer = UK / Double outer = US) --
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblQuoteNesting")
    With lbl
        .Caption = "Outer quotes:"
        .Left = colRight: .Top = yPos + 2: .Width = 80: .Height = LBL_H
    End With

    Set optQuoteSingle = Me.Controls.Add("Forms.OptionButton.1", "optQuoteSingle")
    With optQuoteSingle
        .Caption = "Single"
        .Left = colRight + 82: .Top = yPos: .Width = 60: .Height = CHK_H
        .Value = True
        .GroupName = "QuoteNesting"
    End With

    Set optQuoteDouble = Me.Controls.Add("Forms.OptionButton.1", "optQuoteDouble")
    With optQuoteDouble
        .Caption = "Double"
        .Left = colRight + 144: .Top = yPos: .Width = 60: .Height = CHK_H
        .Value = False
        .GroupName = "QuoteNesting"
    End With
    yPos = yPos + CHK_H + ITEM_GAP

    ' -- Smart quotes toggle (Smart / Straight) --
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblSmartQuotes")
    With lbl
        .Caption = "Smart quotes:"
        .Left = colRight: .Top = yPos + 2: .Width = 80: .Height = LBL_H
    End With

    Set optSmart = Me.Controls.Add("Forms.OptionButton.1", "optSmart")
    With optSmart
        .Caption = "Smart"
        .Left = colRight + 82: .Top = yPos: .Width = 60: .Height = CHK_H
        .Value = True
        .GroupName = "SmartQuotes"
    End With

    Set optSmartStraight = Me.Controls.Add("Forms.OptionButton.1", "optSmartStraight")
    With optSmartStraight
        .Caption = "Straight"
        .Left = colRight + 144: .Top = yPos: .Width = 70: .Height = CHK_H
        .Value = False
        .GroupName = "SmartQuotes"
    End With
    yPos = yPos + CHK_H + ITEM_GAP

    ' -- Date format toggle (UK / US) --
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblDateFormat")
    With lbl
        .Caption = "Date format:"
        .Left = colRight: .Top = yPos + 2: .Width = 80: .Height = LBL_H
    End With

    Set optDateUK = Me.Controls.Add("Forms.OptionButton.1", "optDateUK")
    With optDateUK
        .Caption = "UK"
        .Left = colRight + 82: .Top = yPos: .Width = 50: .Height = CHK_H
        .Value = True
        .GroupName = "DateFormat"
    End With

    Set optDateUS = Me.Controls.Add("Forms.OptionButton.1", "optDateUS")
    With optDateUS
        .Caption = "US"
        .Left = colRight + 134: .Top = yPos: .Width = 50: .Height = CHK_H
        .Value = False
        .GroupName = "DateFormat"
    End With

    yPos = yPos + CHK_H + ITEM_GAP

    ' -- Defined Terms: [format dropdown] and [quotes dropdown] --
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblDefinedTerms")
    With lbl
        .Caption = "Defined Terms:"
        .Left = colRight: .Top = yPos + 2: .Width = 80: .Height = LBL_H
    End With

    Set cboTermFormat = Me.Controls.Add("Forms.ComboBox.1", "cboTermFormat")
    With cboTermFormat
        .Left = colRight + 82: .Top = yPos: .Width = 90: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "Bold"
        .AddItem "Bold Italics"
        .AddItem "Italics"
        .AddItem "None"
        .ListIndex = 0
    End With

    Dim lblAnd As MSForms.Label
    Set lblAnd = Me.Controls.Add("Forms.Label.1", "lblTermAnd")
    With lblAnd
        .Caption = "and"
        .Left = colRight + 175: .Top = yPos + 2: .Width = 22: .Height = LBL_H
    End With

    Set cboTermQuotes = Me.Controls.Add("Forms.ComboBox.1", "cboTermQuotes")
    With cboTermQuotes
        .Left = colRight + 198: .Top = yPos: .Width = 100: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "Single quotes"
        .AddItem "Double quotes"
        .ListIndex = 1
    End With

    yPos = yPos + TXT_H + ITEM_GAP

    ' -- Space style after full stops dropdown --
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblSpaceStyle")
    With lbl
        .Caption = "After full stop:"
        .Left = colRight: .Top = yPos + 2: .Width = 80: .Height = LBL_H
    End With

    Set cboSpaceStyle = Me.Controls.Add("Forms.ComboBox.1", "cboSpaceStyle")
    With cboSpaceStyle
        .Left = colRight + 82: .Top = yPos: .Width = 120: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "One space"
        .AddItem "Two spaces"
        .ListIndex = 0
    End With

    yPos = yPos + TXT_H + SEC_GAP

    ' ==========================================================
    '  ROW 4: Brand Rules
    ' ==========================================================
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblBrandHeader")
    With lbl
        .Caption = "Brand Rules"
        .Left = PAD: .Top = yPos: .Width = 200: .Height = LBL_H
        .Font.Size = 10: .Font.Bold = True
    End With
    yPos = yPos + LBL_H + ITEM_GAP

    Set lstBrands = Me.Controls.Add("Forms.ListBox.1", "lstBrands")
    With lstBrands
        .Left = PAD: .Top = yPos: .Width = FULL_W - BTN_W - SEC_GAP
        .Height = 72
    End With

    ' Brand action buttons (right of list)
    Dim btnX As Single
    btnX = PAD + lstBrands.Width + ITEM_GAP
    Dim brandBtnY As Single
    brandBtnY = yPos

    Set btnAddBrand = Me.Controls.Add("Forms.CommandButton.1", "btnAddBrand")
    With btnAddBrand
        .Caption = "Add"
        .Left = btnX: .Top = brandBtnY: .Width = BTN_W: .Height = BTN_H
    End With
    brandBtnY = brandBtnY + BTN_H + 2

    Set btnRemoveBrand = Me.Controls.Add("Forms.CommandButton.1", "btnRemoveBrand")
    With btnRemoveBrand
        .Caption = "Remove"
        .Left = btnX: .Top = brandBtnY: .Width = BTN_W: .Height = BTN_H
    End With

    ' Save/Load beside Add/Remove
    Dim btnX2 As Single
    btnX2 = btnX
    brandBtnY = brandBtnY + BTN_H + 2

    Set btnSaveBrands = Me.Controls.Add("Forms.CommandButton.1", "btnSaveBrands")
    With btnSaveBrands
        .Caption = "Save Rules"
        .Left = btnX2: .Top = brandBtnY: .Width = BTN_W / 2 - 1: .Height = BTN_H
        .Font.Size = 8
    End With

    Set btnLoadBrands = Me.Controls.Add("Forms.CommandButton.1", "btnLoadBrands")
    With btnLoadBrands
        .Caption = "Load Rules"
        .Left = btnX2 + BTN_W / 2 + 1: .Top = brandBtnY: .Width = BTN_W / 2 - 1: .Height = BTN_H
        .Font.Size = 8
    End With

    yPos = yPos + lstBrands.Height + ITEM_GAP

    ' Brand input fields
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblCorrectForm")
    With lbl
        .Caption = "Correct Form:"
        .Left = PAD: .Top = yPos + 3: .Width = 78: .Height = LBL_H
    End With

    Set txtBrandCorrect = Me.Controls.Add("Forms.TextBox.1", "txtBrandCorrect")
    With txtBrandCorrect
        .Left = PAD + 78: .Top = yPos: .Width = 150: .Height = TXT_H
    End With

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblIncorrectVars")
    With lbl
        .Caption = "Incorrect Variants:"
        .Left = PAD + 240: .Top = yPos + 3: .Width = 108: .Height = LBL_H
    End With

    Set txtBrandIncorrect = Me.Controls.Add("Forms.TextBox.1", "txtBrandIncorrect")
    With txtBrandIncorrect
        .Left = PAD + 348: .Top = yPos: .Width = 180: .Height = TXT_H
    End With

    yPos = yPos + TXT_H + SEC_GAP

    ' ==========================================================
    '  ROW 5: Action Buttons
    ' ==========================================================
    Const ACT_BTN_H As Single = 32
    Const ACT_BTN_W As Single = 120
    Const ACT_GAP As Single = 10

    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    With btnRun
        .Caption = "Run Checks"
        .Left = PAD: .Top = yPos: .Width = ACT_BTN_W + 20: .Height = ACT_BTN_H
        .Font.Bold = True
    End With

    Set btnExport = Me.Controls.Add("Forms.CommandButton.1", "btnExport")
    With btnExport
        .Caption = "Export Report"
        .Left = PAD + ACT_BTN_W + 20 + ACT_GAP: .Top = yPos
        .Width = ACT_BTN_W: .Height = ACT_BTN_H
    End With

    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close"
        .Left = PAD + 3 * (ACT_BTN_W + ACT_GAP) + 12: .Top = yPos
        .Width = 84: .Height = ACT_BTN_H
    End With

    yPos = yPos + ACT_BTN_H + ITEM_GAP

    ' ==========================================================
    '  ROW 6: Status Bar
    ' ==========================================================
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    With lblStatus
        .Caption = "Ready. Select rules and click Run."
        .Left = PAD: .Top = yPos: .Width = FULL_W: .Height = LBL_H
        .Font.Size = 9
    End With

    ' -- Load brand list ---------------------------------------
    RefreshBrandList

    ' -- Hardcoded form size: 1000 x 1000 points ---------------
    ' VBA UserForm Width/Height are in points.
    ' Set explicitly here as a defensive override in case the
    ' .frm persisted ClientWidth/ClientHeight values are ignored
    ' or overridden by control layout.
    Me.Width = 1000
    Me.Height = 1000
    Debug.Print "UserForm_Initialize: Width=" & Me.Width & " Height=" & Me.Height
End Sub

' ============================================================
'  BUILD DYNAMIC RULE CHECKBOX LIST
' ============================================================
Private Sub BuildRuleCheckboxList(nRules As Long)
    Set ruleCheckboxes = New Collection

    Dim chk As MSForms.CheckBox
    Dim displayLabel As String
    Dim i As Long

    ' Multi-column layout: 4 columns across the wide frame
    Const COLS As Long = 4
    Const ROW_H As Single = 18
    Const COL_PAD As Single = 6

    Dim colW As Single
    colW = (fraRules.InsideWidth - COL_PAD * 2) / COLS

    Dim col As Long
    Dim row As Long

    For i = 0 To nRules - 1
        If ruleDisplayMap.Exists(ruleKeys(i)) Then
            displayLabel = CStr(i + 1) & ". " & CStr(ruleDisplayMap(ruleKeys(i)))
        Else
            displayLabel = CStr(i + 1) & ". " & ruleKeys(i)
        End If

        col = i Mod COLS
        row = i \ COLS

        Set chk = fraRules.Controls.Add("Forms.CheckBox.1", "chkRule_" & i)
        With chk
            .Caption = displayLabel
            .Left = COL_PAD + col * colW
            .Top = COL_PAD + row * ROW_H
            .Width = colW - 4
            .Height = ROW_H
            .Value = True
            .Font.Size = 8
        End With

        ruleCheckboxes.Add chk
    Next i

    Dim totalRows As Long
    totalRows = (nRules + COLS - 1) \ COLS
    fraRules.ScrollHeight = COL_PAD * 2 + totalRows * ROW_H
End Sub

' ============================================================
'  RUN BUTTON
' ============================================================
Private Sub btnRun_Click()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Pleadings Checker"
        Exit Sub
    End If

    ' Sync rule config from dynamic checkboxes
    Dim i As Long
    For i = 1 To ruleCheckboxes.Count
        Dim rName As String
        rName = ruleKeys(i - 1)
        If ruleConfig.Exists(rName) Then
            ruleConfig(rName) = CBool(ruleCheckboxes(i).Value)
        End If
    Next i

    ' Set page range from flexible input
    PleadingsEngine.SetPageRangeFromString txtPageRange.Text

    ' Set mode toggles
    If optSpellingUS.Value Then
        PleadingsEngine.SetSpellingMode "US"
    Else
        PleadingsEngine.SetSpellingMode "UK"
    End If

    If optQuoteDouble.Value Then
        PleadingsEngine.SetQuoteNesting "DOUBLE"
    Else
        PleadingsEngine.SetQuoteNesting "SINGLE"
    End If

    If optSmartStraight.Value Then
        PleadingsEngine.SetSmartQuotePref "STRAIGHT"
    Else
        PleadingsEngine.SetSmartQuotePref "SMART"
    End If

    If optDateUS.Value Then
        PleadingsEngine.SetDateFormatPref "US"
    Else
        PleadingsEngine.SetDateFormatPref "UK"
    End If

    ' Set defined term detection preferences
    Dim termFmt As String
    Select Case cboTermFormat.ListIndex
        Case 0: termFmt = "BOLD"
        Case 1: termFmt = "BOLDITALIC"
        Case 2: termFmt = "ITALIC"
        Case Else: termFmt = "NONE"
    End Select
    PleadingsEngine.SetTermFormatPref termFmt

    Dim termQt As String
    If cboTermQuotes.ListIndex = 0 Then
        termQt = "SINGLE"
    Else
        termQt = "DOUBLE"
    End If
    PleadingsEngine.SetTermQuotePref termQt

    ' Set space style preference
    If cboSpaceStyle.ListIndex = 1 Then
        PleadingsEngine.SetSpaceStylePref "TWO"
    Else
        PleadingsEngine.SetSpaceStylePref "ONE"
    End If

    ' Run checks
    lblStatus.Caption = "Running checks..."
    Me.Repaint
    DoEvents

    Set lastResults = PleadingsEngine.RunAllPleadingsRules(ActiveDocument, ruleConfig)

    ' Show performance summary in Immediate window
    Dim slowestRules As String
    If PleadingsEngine.ENABLE_PROFILING Then
        Dim perfSummary As String
        perfSummary = PleadingsEngine.GetPerformanceSummary()
        slowestRules = PleadingsEngine.GetTopSlowestRules(3)
        Debug.Print "UserForm final: Width=" & Me.Width & " Height=" & Me.Height
    End If

    ' Show summary
    Dim summary As String
    summary = PleadingsEngine.GetIssueSummary(lastResults)

    Dim errCount As Long
    errCount = PleadingsEngine.GetRuleErrorCount()

    If lastResults.Count = 0 Then
        If errCount > 0 Then
            lblStatus.Caption = "No issues found, but " & errCount & " rule(s) failed."
            MsgBox "No issues found, but " & errCount & " rule(s) failed to run:" & vbCrLf & vbCrLf & _
                   PleadingsEngine.GetRuleErrorLog() & vbCrLf & _
                   "Check Immediate window (Ctrl+G) or export a report for the debug log.", _
                   vbExclamation, "Pleadings Checker"
        Else
            lblStatus.Caption = "No issues found."
            MsgBox "No issues found -- document looks clean.", vbInformation, "Pleadings Checker"
        End If
    Else
        Dim errMsg As String
        If errCount > 0 Then
            errMsg = vbCrLf & errCount & " rule(s) failed to run:" & vbCrLf & _
                     PleadingsEngine.GetRuleErrorLog()
        End If
        If Len(slowestRules) > 0 Then
            errMsg = errMsg & vbCrLf & "Slowest: " & slowestRules
        End If

        Dim reply As VbMsgBoxResult
        reply = MsgBox(lastResults.Count & " issue(s) found." & errMsg & vbCrLf & vbCrLf & _
               "Apply suggestions to the document?", _
               vbYesNo + vbQuestion, "Pleadings Checker")

        If reply = vbYes Then
            lblStatus.Caption = "Applying suggestions..."
            Me.Repaint
            DoEvents

            Dim addComments As Boolean
            addComments = (chkAddComments.Value = True)

            If chkTrackedChanges.Value = True Then
                PleadingsEngine.ApplySuggestionsAsTrackedChanges ActiveDocument, lastResults, addComments
            Else
                PleadingsEngine.ApplyHighlights ActiveDocument, lastResults, addComments
            End If

            lblStatus.Caption = lastResults.Count & " issue(s) applied."
        Else
            lblStatus.Caption = lastResults.Count & " issue(s) found. Use Export Report for details."
        End If
    End If
End Sub

' ============================================================
'  EXPORT REPORT BUTTON
' ============================================================
Private Sub btnExport_Click()
    If lastResults Is Nothing Then
        MsgBox "Run checks first before exporting.", vbExclamation, "Pleadings Checker"
        Exit Sub
    End If

    Dim reportPath As String
    Dim sep As String
    sep = Application.PathSeparator

    On Error Resume Next
    If ActiveDocument.Path <> "" Then
        Dim baseName As String
        baseName = ActiveDocument.Name
        Dim dotPos As Long
        dotPos = InStrRev(baseName, ".")
        If dotPos > 1 Then baseName = Left$(baseName, dotPos - 1)
        reportPath = ActiveDocument.Path & sep & baseName & "_pleadings_report.json"
    End If
    If Err.Number <> 0 Or Len(reportPath) = 0 Then
        Err.Clear
        reportPath = ""
    End If
    On Error GoTo 0

    ' Fallback to temp directory if no valid path yet
    If Len(reportPath) = 0 Then
        reportPath = GetTempReportPath(sep)
    End If

    ' Ensure parent directory exists before writing
    Dim reportDir As String
    reportDir = modDebugLog.GetParentDirectory(reportPath)
    If Len(reportDir) > 0 Then
        modDebugLog.EnsureDirectoryExists reportDir
    End If

    lblStatus.Caption = "Exporting report..."
    Me.Repaint
    DoEvents

    Dim summary As String
    summary = PleadingsEngine.GenerateReport(lastResults, reportPath, ActiveDocument)

    ' Auto-save debug log alongside report when DEBUG_MODE is True
    Dim logPath As String
    Dim logSaved As Boolean
    logSaved = False
    logPath = ""

    On Error Resume Next
    If modDebugLog.DEBUG_MODE Then
        logPath = Left$(reportPath, Len(reportPath) - 5) & "_debug.log"
        logSaved = modDebugLog.DebugLogSaveToTextFile(logPath)
    End If
    On Error GoTo 0

    ' Build informative export message
    Dim errCount As Long
    errCount = PleadingsEngine.GetRuleErrorCount()

    Dim msg As String
    msg = "Report saved to:" & vbCrLf & reportPath

    If logSaved And Len(logPath) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Debug log saved to:" & vbCrLf & logPath
    ElseIf modDebugLog.DEBUG_MODE And Not logSaved Then
        msg = msg & vbCrLf & vbCrLf & "Debug log could not be saved."
    End If

    If errCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & errCount & " rule(s) failed during the run."
    End If

    msg = msg & vbCrLf & vbCrLf & summary

    lblStatus.Caption = "Report saved."
    MsgBox msg, vbInformation, "Pleadings Checker -- Report"
End Sub

' -- Helper: build a temp path for report export (cross-platform) --
Private Function GetTempReportPath(sep As String) As String
    Dim tmpDir As String
    #If Mac Then
        tmpDir = Environ("TMPDIR")
        If Len(tmpDir) = 0 Then tmpDir = "/tmp"
        ' Strip trailing separator if present
        If Right$(tmpDir, 1) = sep Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
    #Else
        tmpDir = Environ("TEMP")
        If Len(tmpDir) = 0 Then tmpDir = Environ("TMP")
        If Len(tmpDir) = 0 Then tmpDir = Environ("USERPROFILE")
        If Len(tmpDir) = 0 Then tmpDir = "C:\Temp"
        If Right$(tmpDir, 1) = sep Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
    #End If
    GetTempReportPath = tmpDir & sep & "pleadings_report.json"
End Function

' ============================================================
'  SELECT ALL / DESELECT ALL
' ============================================================
Private Sub btnSelectAll_Click()
    Dim i As Long
    For i = 1 To ruleCheckboxes.Count
        ruleCheckboxes(i).Value = True
    Next i
End Sub

Private Sub btnDeselectAll_Click()
    Dim i As Long
    For i = 1 To ruleCheckboxes.Count
        ruleCheckboxes(i).Value = False
    Next i
End Sub

' ============================================================
'  BRAND RULES MANAGEMENT
' ============================================================
Private Sub RefreshBrandList()
    lstBrands.Clear
    On Error Resume Next
    Dim brands As Object
    Set brands = Application.Run("Rules_Brands.GetBrandRules")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
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

    On Error Resume Next
    Application.Run "Rules_Brands.AddBrandRule", correctForm, incorrectVariants
    If Err.Number <> 0 Then
        MsgBox "Brand rules module not loaded.", vbExclamation, "Brand Rules"
        Err.Clear
    End If
    On Error GoTo 0

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
    Dim correctForm As String
    correctForm = Left(entry, InStr(entry, " -> ") - 1)

    On Error Resume Next
    Application.Run "Rules_Brands.RemoveBrandRule", correctForm
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0

    RefreshBrandList
End Sub

Private Sub btnSaveBrands_Click()
    Dim brandFile As String
    brandFile = GetBrandRulesPath()

    ' Ensure directory exists (recursive, handles nested paths)
    Dim brandDir As String
    brandDir = modDebugLog.GetParentDirectory(brandFile)
    If Len(brandDir) > 0 Then
        modDebugLog.EnsureDirectoryExists brandDir
    End If

    Dim saveResult As Boolean
    On Error Resume Next
    saveResult = Application.Run("Rules_Brands.SaveBrandRules", brandFile)
    If Err.Number <> 0 Then
        MsgBox "Brand rules module not loaded." & vbCrLf & _
               "Error: " & Err.Description, vbExclamation, "Brand Rules"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If saveResult Then
        MsgBox "Brand rules saved to:" & vbCrLf & brandFile, vbInformation, "Brand Rules"
    Else
        MsgBox "Failed to save brand rules to:" & vbCrLf & brandFile & vbCrLf & _
               "Check the file path is writable.", vbExclamation, "Brand Rules"
    End If
End Sub

Private Sub btnLoadBrands_Click()
    Dim brandFile As String
    brandFile = GetBrandRulesPath()

    If Dir(brandFile) = "" Then
        MsgBox "No saved brand rules found at:" & vbCrLf & brandFile, _
               vbExclamation, "Brand Rules"
        Exit Sub
    End If

    Dim loadResult As Boolean
    On Error Resume Next
    loadResult = Application.Run("Rules_Brands.LoadBrandRules", brandFile)
    If Err.Number <> 0 Then
        MsgBox "Brand rules module not loaded." & vbCrLf & _
               "Error: " & Err.Description, vbExclamation, "Brand Rules"
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    RefreshBrandList
    If loadResult Then
        MsgBox "Brand rules loaded.", vbInformation, "Brand Rules"
    Else
        MsgBox "Brand rules file could not be read:" & vbCrLf & brandFile, _
               vbExclamation, "Brand Rules"
    End If
End Sub

' -- Helper: cross-platform brand rules file path --
' Delegates to Rules_Brands.GetDefaultBrandRulesPath (single source of truth).
' Falls back to a local construction if the module is not imported.
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

' ============================================================
'  CLOSE BUTTON
' ============================================================
Private Sub btnClose_Click()
    Unload Me
End Sub

```

# FILE: modDebugLog.bas

```vb
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

```

# FILE: PleadingsEngine.bas

```vb
Attribute VB_Name = "PleadingsEngine"
' ============================================================
' PleadingsEngine.bas
' Core engine for the Pleadings Checker rule system.
'
' MODULAR ARCHITECTURE: Uses Application.Run to dispatch rules
' so that missing modules produce trappable runtime errors
' instead of compile errors. Import only the rule modules you
' need -- the engine gracefully skips any that are absent.
'
' Dependencies:
'
' Optional rule modules (import any subset):
'   - Rules_Spelling.bas        (Rules 1, 12, 13)
'   - Rules_TextScan.bas        (Rules 2, 34)
'   - Rules_Numbering.bas       (Rules 3, 8)
'   - Rules_Headings.bas        (Rules 4, 21)
'   - Rules_Terms.bas           (Rules 5, 7; 23 RETIRED)
'   - Rules_Formatting.bas      (Rules 6, 11)
'   - Rules_NumberFormats.bas    (Rules 9, 19; 18 RETIRED)
'   - Rules_Lists.bas           (Rules 10, 15)
'   - Rules_Punctuation.bas     (Rules 14, 16)
'   - Rules_Quotes.bas          (Rules 17, 32, 33)
'   - Rules_FootnoteIntegrity.bas (Rule 20)
'   - Rules_Brands.bas          (Rule 22)
'   - Rules_FootnoteHarts.bas   (Rules 24, 25, 26, 27)
'   - Rules_LegalTerms.bas      (Rules 28, 29)
'   - Rules_Italics.bas         (Rules 30, 31)
'   - Rules_Spacing.bas        (Rules 35-39: double spaces, commas, spacing)
'
' Installation:
'   1. Open the VBA Editor (Alt+F11)
'   2. File > Import File > PleadingsEngine.bas
'   3. File > Import File > PleadingsLauncher.bas
'   4. Import whichever Rules_*.bas modules you need
'   5. Run the macro "PleadingsChecker"
'
'   Note: No early-bound references are required. All Scripting.Dictionary
'   usage is late-bound via CreateObject("Scripting.Dictionary").
' ============================================================
Option Explicit

' -- Module-level state --
Private ruleConfig      As Object
Private pageRangeSet    As Object   ' Dictionary of page numbers (Long -> True)
Private whitelistDict   As Object
Private spellingMode    As String   ' "UK" or "US"
Private quoteNesting   As String   ' "SINGLE" or "DOUBLE" (outer marks)
Private smartQuotePref As String   ' "SMART" or "STRAIGHT"
Private dateFormatPref As String   ' "UK" or "US" or "AUTO"
Private termFormatPref As String   ' "BOLD", "BOLDITALIC", "ITALIC", or "NONE"
Private termQuotePref  As String   ' "SINGLE" or "DOUBLE"
Private spaceStylePref As String   ' "ONE" or "TWO"
Private ruleErrorCount  As Long
Private ruleErrorLog    As String

' -- Profiling infrastructure --
Public Const ENABLE_PROFILING As Boolean = True
Private perfTimings     As Object   ' Dictionary: label -> elapsed Single
Private perfCounters    As Object   ' Dictionary: label -> Long count
Private perfStarts      As Object   ' Dictionary: label -> start Timer value
Private totalStartTime  As Single

' -- Paragraph position cache (built once per run for O(log N) lookups) --
Private paraStartPos()  As Long
Private paraStartCount  As Long
Private paraCacheValid  As Boolean

' ============================================================
'  ENTRY POINT
' ============================================================
Public Sub PleadingsChecker()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Pleadings Checker"
        Exit Sub
    End If
    ' Show the UserForm; fall back to quick run if form not imported
    On Error Resume Next
    frmPleadingsChecker.Show
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        RunQuick
    End If
    On Error GoTo 0
End Sub

' ============================================================
'  QUICK RUN (fallback when launcher is not imported)
'  Runs all available rules and shows summary via MsgBox.
' ============================================================
Public Sub RunQuick()
    TraceEnter "RunQuick"
    DebugLogDoc "RunQuick target", ActiveDocument
    Dim cfg As Object
    Set cfg = InitRuleConfig()
    SetPageRange 0, 0
    SetSpellingMode "UK"

    Dim issues As Collection
    Set issues = RunAllPleadingsRules(ActiveDocument, cfg)

    Dim summary As String
    summary = GetIssueSummary(issues)

    If issues.Count = 0 Then
        MsgBox "No issues found.", vbInformation, "Pleadings Checker"
    Else
        MsgBox summary, vbInformation, "Pleadings Checker"
        ApplySuggestionsAsTrackedChanges ActiveDocument, issues, True
    End If
    TraceExit "RunQuick", issues.Count & " issues"
End Sub

' ============================================================
'  SPELLING MODE (UK / US toggle)
' ============================================================
Public Sub SetSpellingMode(ByVal mode As String)
    spellingMode = UCase(Trim(mode))
    If spellingMode <> "US" Then spellingMode = "UK"
End Sub

Public Function GetSpellingMode() As String
    If Len(spellingMode) = 0 Then spellingMode = "UK"
    GetSpellingMode = spellingMode
End Function

' ============================================================
'  QUOTE NESTING (single outer = UK, double outer = US)
' ============================================================
Public Sub SetQuoteNesting(ByVal mode As String)
    quoteNesting = UCase(Trim(mode))
    If quoteNesting <> "DOUBLE" Then quoteNesting = "SINGLE"
End Sub

Public Function GetQuoteNesting() As String
    If Len(quoteNesting) = 0 Then quoteNesting = "SINGLE"
    GetQuoteNesting = quoteNesting
End Function

' ============================================================
'  SMART QUOTE PREFERENCE (smart or straight)
' ============================================================
Public Sub SetSmartQuotePref(ByVal mode As String)
    smartQuotePref = UCase(Trim(mode))
    If smartQuotePref <> "STRAIGHT" Then smartQuotePref = "SMART"
End Sub

Public Function GetSmartQuotePref() As String
    If Len(smartQuotePref) = 0 Then smartQuotePref = "SMART"
    GetSmartQuotePref = smartQuotePref
End Function

' ============================================================
'  DATE FORMAT PREFERENCE (UK = "1 January 2024", US = "January 1, 2024")
' ============================================================
Public Sub SetDateFormatPref(ByVal mode As String)
    dateFormatPref = UCase(Trim(mode))
    If dateFormatPref <> "US" And dateFormatPref <> "AUTO" Then dateFormatPref = "UK"
End Sub

Public Function GetDateFormatPref() As String
    If Len(dateFormatPref) = 0 Then dateFormatPref = "UK"
    GetDateFormatPref = dateFormatPref
End Function

' ============================================================
'  DEFINED TERM FORMAT PREFERENCE
' ============================================================
Public Sub SetTermFormatPref(ByVal mode As String)
    termFormatPref = UCase(Trim(mode))
    If termFormatPref <> "BOLDITALIC" And termFormatPref <> "ITALIC" And _
       termFormatPref <> "NONE" Then termFormatPref = "BOLD"
End Sub

Public Function GetTermFormatPref() As String
    If Len(termFormatPref) = 0 Then termFormatPref = "BOLD"
    GetTermFormatPref = termFormatPref
End Function

' ============================================================
'  DEFINED TERM QUOTE PREFERENCE
' ============================================================
Public Sub SetTermQuotePref(ByVal mode As String)
    termQuotePref = UCase(Trim(mode))
    If termQuotePref <> "SINGLE" Then termQuotePref = "DOUBLE"
End Sub

Public Function GetTermQuotePref() As String
    If Len(termQuotePref) = 0 Then termQuotePref = "DOUBLE"
    GetTermQuotePref = termQuotePref
End Function

' ============================================================
'  SPACE STYLE PREFERENCE (one space or two after full stop)
' ============================================================
Public Sub SetSpaceStylePref(ByVal mode As String)
    spaceStylePref = UCase(Trim(mode))
    If spaceStylePref <> "TWO" Then spaceStylePref = "ONE"
End Sub

Public Function GetSpaceStylePref() As String
    If Len(spaceStylePref) = 0 Then spaceStylePref = "ONE"
    GetSpaceStylePref = spaceStylePref
End Function

' ============================================================
'  RULE CONFIGURATION
' ============================================================
Public Function InitRuleConfig() As Object
    Dim cfg As Object
    Set cfg = CreateObject("Scripting.Dictionary")

    cfg.Add "spelling", True
    cfg.Add "repeated_words", True
    cfg.Add "sequential_numbering", True
    cfg.Add "heading_capitalisation", True
    cfg.Add "custom_term_whitelist", True
    cfg.Add "defined_terms", True
    cfg.Add "clause_number_format", True
    cfg.Add "date_time_format", True
    cfg.Add "list_rules", True
    cfg.Add "formatting_consistency", True
    cfg.Add "licence_license", True
    cfg.Add "check_cheque", True
    cfg.Add "slash_style", True
    cfg.Add "dash_usage", True
    cfg.Add "bracket_integrity", True
    cfg.Add "quotation_mark_consistency", True
    cfg.Add "currency_number_format", True
    cfg.Add "footnote_rules", True
    cfg.Add "title_formatting", True
    cfg.Add "brand_name_enforcement", True
    cfg.Add "mandated_legal_term_forms", True
    cfg.Add "always_capitalise_terms", True
    cfg.Add "known_anglicised_terms_not_italic", True
    cfg.Add "foreign_names_not_italic", True
    cfg.Add "single_quotes_default", True
    cfg.Add "smart_quote_consistency", True
    cfg.Add "spell_out_under_ten", True
    cfg.Add "double_spaces", True
    cfg.Add "double_commas", True
    cfg.Add "space_before_punct", True
    cfg.Add "missing_space_after_dot", True
    cfg.Add "trailing_spaces", True

    Set InitRuleConfig = cfg
End Function

' ============================================================
'  PROFILING INFRASTRUCTURE
' ============================================================
Public Sub PerfTimerStart(ByVal label As String)
    If Not ENABLE_PROFILING Then Exit Sub
    On Error Resume Next
    If perfStarts Is Nothing Then Set perfStarts = CreateObject("Scripting.Dictionary")
    perfStarts(label) = Timer
    On Error GoTo 0
End Sub

Public Sub PerfTimerEnd(ByVal label As String)
    If Not ENABLE_PROFILING Then Exit Sub
    On Error Resume Next
    If perfTimings Is Nothing Then Set perfTimings = CreateObject("Scripting.Dictionary")
    Dim elapsed As Single
    elapsed = Timer - CSng(perfStarts(label))
    If elapsed < 0 Then elapsed = elapsed + 86400  ' midnight rollover
    If perfTimings.Exists(label) Then
        perfTimings(label) = CSng(perfTimings(label)) + elapsed
    Else
        perfTimings(label) = elapsed
    End If
    On Error GoTo 0
End Sub

Public Sub PerfCount(ByVal label As String, Optional ByVal increment As Long = 1)
    If Not ENABLE_PROFILING Then Exit Sub
    On Error Resume Next
    If perfCounters Is Nothing Then Set perfCounters = CreateObject("Scripting.Dictionary")
    If perfCounters.Exists(label) Then
        perfCounters(label) = CLng(perfCounters(label)) + increment
    Else
        perfCounters(label) = increment
    End If
    On Error GoTo 0
End Sub

Private Sub ResetProfiling()
    Set perfTimings = CreateObject("Scripting.Dictionary")
    Set perfCounters = CreateObject("Scripting.Dictionary")
    Set perfStarts = CreateObject("Scripting.Dictionary")
    totalStartTime = Timer
    paraCacheValid = False
End Sub

Public Function GetPerformanceSummary() As String
    If Not ENABLE_PROFILING Then
        GetPerformanceSummary = "(Profiling disabled)"
        Exit Function
    End If

    Dim totalElapsed As Single
    totalElapsed = Timer - totalStartTime
    If totalElapsed < 0 Then totalElapsed = totalElapsed + 86400

    Dim result As String
    result = "=== PERFORMANCE SUMMARY ===" & vbCrLf
    result = result & "Total runtime: " & Format(totalElapsed, "0.00") & "s" & vbCrLf & vbCrLf

    ' Sort timings by descending elapsed time
    If Not perfTimings Is Nothing Then
        If perfTimings.Count > 0 Then
            Dim labels() As String
            Dim times() As Single
            Dim n As Long
            n = perfTimings.Count
            ReDim labels(0 To n - 1)
            ReDim times(0 To n - 1)
            Dim keys As Variant
            keys = perfTimings.keys
            Dim idx As Long
            For idx = 0 To n - 1
                labels(idx) = CStr(keys(idx))
                times(idx) = CSng(perfTimings(keys(idx)))
            Next idx

            ' Bubble sort descending by time (small N)
            Dim swapped As Boolean
            Dim tmpS As String
            Dim tmpF As Single
            Do
                swapped = False
                Dim si As Long
                For si = 0 To n - 2
                    If times(si) < times(si + 1) Then
                        tmpF = times(si): times(si) = times(si + 1): times(si + 1) = tmpF
                        tmpS = labels(si): labels(si) = labels(si + 1): labels(si + 1) = tmpS
                        swapped = True
                    End If
                Next si
            Loop While swapped

            result = result & "-- Timings (slowest first) --" & vbCrLf
            For idx = 0 To n - 1
                result = result & "  " & labels(idx) & ": " & Format(times(idx), "0.000") & "s"
                If Not perfCounters Is Nothing Then
                    If perfCounters.Exists(labels(idx) & "_count") Then
                        result = result & " (" & perfCounters(labels(idx) & "_count") & " items)"
                    End If
                End If
                result = result & vbCrLf
            Next idx
        End If
    End If

    ' Counters section
    If Not perfCounters Is Nothing Then
        If perfCounters.Count > 0 Then
            result = result & vbCrLf & "-- Counters --" & vbCrLf
            keys = perfCounters.keys
            For idx = 0 To perfCounters.Count - 1
                result = result & "  " & CStr(keys(idx)) & ": " & perfCounters(keys(idx)) & vbCrLf
            Next idx
        End If
    End If

    GetPerformanceSummary = result
End Function

Public Function GetTopSlowestRules(Optional ByVal topN As Long = 3) As String
    If Not ENABLE_PROFILING Then
        GetTopSlowestRules = ""
        Exit Function
    End If
    If perfTimings Is Nothing Then Exit Function
    If perfTimings.Count = 0 Then Exit Function

    ' Build sorted arrays (same sort as GetPerformanceSummary)
    Dim labels() As String, times() As Single
    Dim n As Long
    n = perfTimings.Count
    ReDim labels(0 To n - 1)
    ReDim times(0 To n - 1)
    Dim keys As Variant
    keys = perfTimings.keys
    Dim idx As Long
    For idx = 0 To n - 1
        labels(idx) = CStr(keys(idx))
        times(idx) = CSng(perfTimings(keys(idx)))
    Next idx

    ' Bubble sort descending
    Dim swapped As Boolean
    Dim tmpS As String: Dim tmpF As Single
    Do
        swapped = False
        Dim si As Long
        For si = 0 To n - 2
            If times(si) < times(si + 1) Then
                tmpF = times(si): times(si) = times(si + 1): times(si + 1) = tmpF
                tmpS = labels(si): labels(si) = labels(si + 1): labels(si + 1) = tmpS
                swapped = True
            End If
        Next si
    Loop While swapped

    Dim result As String
    Dim limit As Long
    limit = topN
    If limit > n Then limit = n
    For idx = 0 To limit - 1
        If idx > 0 Then result = result & ", "
        result = result & labels(idx) & " (" & Format(times(idx), "0.0") & "s)"
    Next idx
    GetTopSlowestRules = result
End Function

' ============================================================
'  PARAGRAPH CACHE (built once per run for O(log N) lookups)
' ============================================================
Private Sub BuildParagraphCache(doc As Document)
    If paraCacheValid Then Exit Sub
    PerfTimerStart "BuildParagraphCache"

    Dim para As Paragraph
    Dim cap As Long
    cap = 512
    ReDim paraStartPos(0 To cap - 1)
    paraStartCount = 0

    On Error Resume Next
    For Each para In doc.Paragraphs
        If paraStartCount >= cap Then
            cap = cap * 2
            ReDim Preserve paraStartPos(0 To cap - 1)
        End If
        paraStartPos(paraStartCount) = para.Range.Start
        paraStartCount = paraStartCount + 1
    Next para
    On Error GoTo 0

    paraCacheValid = True
    PerfTimerEnd "BuildParagraphCache"
    PerfCount "paragraphs_cached", paraStartCount
End Sub

Private Function FindParagraphIndex(ByVal pos As Long) As Long
    If Not paraCacheValid Or paraStartCount = 0 Then
        FindParagraphIndex = 0
        Exit Function
    End If

    ' Binary search for paragraph containing this position
    Dim lo As Long, hi As Long, mid As Long
    lo = 0
    hi = paraStartCount - 1

    Do While lo <= hi
        mid = (lo + hi) \ 2
        If mid < paraStartCount - 1 Then
            If paraStartPos(mid) <= pos And paraStartPos(mid + 1) > pos Then
                FindParagraphIndex = mid + 1  ' 1-based
                Exit Function
            ElseIf paraStartPos(mid) > pos Then
                hi = mid - 1
            Else
                lo = mid + 1
            End If
        Else
            ' Last paragraph
            If paraStartPos(mid) <= pos Then
                FindParagraphIndex = mid + 1
            Else
                FindParagraphIndex = mid
            End If
            Exit Function
        End If
    Loop

    FindParagraphIndex = lo + 1  ' 1-based
End Function

' ============================================================
'  APPLICATION.RUN DISPATCHER
'  Calls a public function by string name. Returns a
'  Collection of issue dictionary, or an empty Collection if
'  the module/function is not available.
' ============================================================
Private Function TryRunRule(ByVal funcName As String, _
                             ByVal doc As Document) As Collection
    Dim result As Object
    Set result = Nothing

    TraceStep "RunAllPleadingsRules", "dispatching " & funcName

    On Error Resume Next
    Set result = Application.Run(funcName, doc)
    If Err.Number <> 0 Then
        ruleErrorCount = ruleErrorCount + 1
        ruleErrorLog = ruleErrorLog & funcName & " (Err " & Err.Number & ": " & Err.Description & ")" & vbCrLf
        DebugLogError "TryRunRule", funcName, Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        Set TryRunRule = New Collection
        Exit Function
    End If
    On Error GoTo 0

    If result Is Nothing Then
        TraceStep "TryRunRule", funcName & " -> 0 issues (Nothing)"
        Set TryRunRule = New Collection
    Else
        TraceStep "TryRunRule", funcName & " -> " & result.Count & " issue(s)"
        Set TryRunRule = result
    End If
End Function

' ============================================================
'  MASTER RULE RUNNER
' ============================================================
Public Function RunAllPleadingsRules(doc As Document, _
                                     config As Object) As Collection
    TraceEnter "RunAllPleadingsRules"
    DebugLogDoc "RunAllPleadingsRules target", doc

    Dim allIssues As New Collection
    Set ruleConfig = config
    ruleErrorCount = 0
    ruleErrorLog = ""

    ' -- Initialise profiling --
    ResetProfiling

    ' -- Capture and suppress screen redraws for performance ----
    Dim wasScreenUpdating As Boolean
    wasScreenUpdating = Application.ScreenUpdating
    Dim wasStatusBar As Variant
    wasStatusBar = Application.StatusBar
    Application.ScreenUpdating = False

    On Error GoTo RunnerCleanup

    ' -- Build paragraph position cache (one scan, enables O(log N) lookups) --
    BuildParagraphCache doc

    ' -- Whitelist rule first (populates whitelistDict) --
    If IsRuleEnabled(config, "custom_term_whitelist") Then
        PerfTimerStart "custom_term_whitelist"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Terms.Check_CustomTermWhitelist", doc)
        PerfTimerEnd "custom_term_whitelist"
    End If
    DoEvents

    ' -- Spelling (bidirectional UK/US) --
    If IsRuleEnabled(config, "spelling") Then
        PerfTimerStart "spelling"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_Spelling", doc)
        PerfTimerEnd "spelling"
    End If

    DoEvents
    ' -- Text scanning rules --
    If IsRuleEnabled(config, "repeated_words") Then
        PerfTimerStart "repeated_words"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_TextScan.Check_RepeatedWords", doc)
        PerfTimerEnd "repeated_words"
    End If

    If IsRuleEnabled(config, "spell_out_under_ten") Then
        PerfTimerStart "spell_out_under_ten"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_TextScan.Check_SpellOutUnderTen", doc)
        PerfTimerEnd "spell_out_under_ten"
    End If

    DoEvents
    ' -- Spacing rules --
    If IsRuleEnabled(config, "double_spaces") Then
        PerfTimerStart "double_spaces"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spacing.Check_DoubleSpaces", doc)
        PerfTimerEnd "double_spaces"
    End If

    If IsRuleEnabled(config, "double_commas") Then
        PerfTimerStart "double_commas"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spacing.Check_DoubleCommas", doc)
        PerfTimerEnd "double_commas"
    End If

    If IsRuleEnabled(config, "space_before_punct") Then
        PerfTimerStart "space_before_punct"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spacing.Check_SpaceBeforePunct", doc)
        PerfTimerEnd "space_before_punct"
    End If

    If IsRuleEnabled(config, "missing_space_after_dot") Then
        PerfTimerStart "missing_space_after_dot"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spacing.Check_MissingSpaceAfterDot", doc)
        PerfTimerEnd "missing_space_after_dot"
    End If

    If IsRuleEnabled(config, "trailing_spaces") Then
        PerfTimerStart "trailing_spaces"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spacing.Check_TrailingSpaces", doc)
        PerfTimerEnd "trailing_spaces"
    End If

    DoEvents
    ' -- Numbering rules --
    If IsRuleEnabled(config, "sequential_numbering") Then
        PerfTimerStart "sequential_numbering"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Numbering.Check_SequentialNumbering", doc)
        PerfTimerEnd "sequential_numbering"
    End If

    If IsRuleEnabled(config, "clause_number_format") Then
        PerfTimerStart "clause_number_format"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Numbering.Check_ClauseNumberFormat", doc)
        PerfTimerEnd "clause_number_format"
    End If

    DoEvents
    ' -- Heading rules --
    If IsRuleEnabled(config, "heading_capitalisation") Then
        PerfTimerStart "heading_capitalisation"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Headings.Check_HeadingCapitalisation", doc)
        PerfTimerEnd "heading_capitalisation"
    End If

    If IsRuleEnabled(config, "title_formatting") Then
        PerfTimerStart "title_formatting"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Headings.Check_TitleFormatting", doc)
        PerfTimerEnd "title_formatting"
    End If

    DoEvents
    ' -- Term rules --
    If IsRuleEnabled(config, "defined_terms") Then
        PerfTimerStart "defined_terms"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Terms.Check_DefinedTerms", doc)
        PerfTimerEnd "defined_terms"
    End If

    DoEvents
    ' -- Formatting consistency (combined: paragraph breaks, font, colour) --
    If IsRuleEnabled(config, "formatting_consistency") Then
        PerfTimerStart "formatting_consistency"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Formatting.Check_ParagraphBreakConsistency", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Formatting.Check_FontConsistency", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_ColourFormatting", doc)
        PerfTimerEnd "formatting_consistency"
    End If

    DoEvents
    ' -- Number format rules --
    If IsRuleEnabled(config, "date_time_format") Then
        PerfTimerStart "date_time_format"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_NumberFormats.Check_DateTimeFormat", doc)
        PerfTimerEnd "date_time_format"
    End If

    If IsRuleEnabled(config, "currency_number_format") Then
        PerfTimerStart "currency_number_format"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_NumberFormats.Check_CurrencyNumberFormat", doc)
        PerfTimerEnd "currency_number_format"
    End If

    DoEvents
    ' -- List rules (combined: inline format, punctuation) --
    If IsRuleEnabled(config, "list_rules") Then
        PerfTimerStart "list_rules"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Lists.Check_InlineListFormat", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Lists.Check_ListPunctuation", doc)
        PerfTimerEnd "list_rules"
    End If

    DoEvents
    ' -- UK/US variant rules (in Rules_Spelling) --
    If IsRuleEnabled(config, "licence_license") Then
        PerfTimerStart "licence_license"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_LicenceLicense", doc)
        PerfTimerEnd "licence_license"
    End If

    If IsRuleEnabled(config, "check_cheque") Then
        PerfTimerStart "check_cheque"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_CheckCheque", doc)
        PerfTimerEnd "check_cheque"
    End If

    DoEvents
    ' -- Punctuation rules --
    If IsRuleEnabled(config, "slash_style") Then
        PerfTimerStart "slash_style"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Punctuation.Check_SlashStyle", doc)
        PerfTimerEnd "slash_style"
    End If

    If IsRuleEnabled(config, "bracket_integrity") Then
        PerfTimerStart "bracket_integrity"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Punctuation.Check_BracketIntegrity", doc)
        PerfTimerEnd "bracket_integrity"
    End If

    If IsRuleEnabled(config, "dash_usage") Then
        PerfTimerStart "dash_usage"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Punctuation.Check_DashUsage", doc)
        PerfTimerEnd "dash_usage"
    End If

    DoEvents
    ' -- Quote rules --
    If IsRuleEnabled(config, "quotation_mark_consistency") Then
        PerfTimerStart "quotation_mark_consistency"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Quotes.Check_QuotationMarkConsistency", doc)
        PerfTimerEnd "quotation_mark_consistency"
    End If

    If IsRuleEnabled(config, "single_quotes_default") Then
        PerfTimerStart "single_quotes_default"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Quotes.Check_SingleQuotesDefault", doc)
        PerfTimerEnd "single_quotes_default"
    End If

    If IsRuleEnabled(config, "smart_quote_consistency") Then
        PerfTimerStart "smart_quote_consistency"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Quotes.Check_SmartQuoteConsistency", doc)
        PerfTimerEnd "smart_quote_consistency"
    End If

    ' -- Dedupe overlapping quote-rule findings --
    ' The three quote rules can flag the same character position independently.
    ' Keep the first finding per RangeStart+RangeEnd and discard later duplicates.
    If allIssues.Count > 0 Then
        Dim quoteRules As Object
        Set quoteRules = CreateObject("Scripting.Dictionary")
        quoteRules.Add "quotation_mark_consistency", True
        quoteRules.Add "single_quotes_default", True
        quoteRules.Add "smart_quote_consistency", True

        Dim seenQuoteKeys As Object
        Set seenQuoteKeys = CreateObject("Scripting.Dictionary")
        Dim dedupedIssues As New Collection
        Dim iss As Variant
        Dim posKey As String

        For Each iss In allIssues
            If quoteRules.Exists(GetIssueProp(iss, "RuleName")) Then
                posKey = GetIssueProp(iss, "RangeStart") & "|" & _
                         GetIssueProp(iss, "RangeEnd")
                If Not seenQuoteKeys.Exists(posKey) Then
                    seenQuoteKeys.Add posKey, GetIssueProp(iss, "RuleName")
                    dedupedIssues.Add iss
                End If
                ' If posKey already seen (from a different quote rule), skip
            Else
                dedupedIssues.Add iss
            End If
        Next iss

        Set allIssues = dedupedIssues
        Set seenQuoteKeys = Nothing
        Set quoteRules = Nothing
    End If

    DoEvents
    ' -- Footnote rules (combined: integrity, not-endnotes, Hart's rules) --
    If IsRuleEnabled(config, "footnote_rules") Then
        PerfTimerStart "footnote_rules"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_FootnoteIntegrity.Check_FootnoteIntegrity", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_FootnoteHarts.Check_FootnotesNotEndnotes", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_FootnoteHarts.Check_FootnoteTerminalFullStop", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_FootnoteHarts.Check_FootnoteInitialCapital", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_FootnoteHarts.Check_FootnoteAbbreviationDictionary", doc)
        PerfTimerEnd "footnote_rules"
    End If

    DoEvents
    ' -- Brand names --
    If IsRuleEnabled(config, "brand_name_enforcement") Then
        PerfTimerStart "brand_name_enforcement"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Brands.Check_BrandNameEnforcement", doc)
        PerfTimerEnd "brand_name_enforcement"
    End If

    DoEvents
    ' -- Legal term rules --
    If IsRuleEnabled(config, "mandated_legal_term_forms") Then
        PerfTimerStart "mandated_legal_term_forms"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_LegalTerms.Check_MandatedLegalTermForms", doc)
        PerfTimerEnd "mandated_legal_term_forms"
    End If

    If IsRuleEnabled(config, "always_capitalise_terms") Then
        PerfTimerStart "always_capitalise_terms"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_LegalTerms.Check_AlwaysCapitaliseTerms", doc)
        PerfTimerEnd "always_capitalise_terms"
    End If

    DoEvents
    ' -- Italic rules --
    If IsRuleEnabled(config, "known_anglicised_terms_not_italic") Then
        PerfTimerStart "anglicised_terms_not_italic"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Italics.Check_AnglicisedTermsNotItalic", doc)
        PerfTimerEnd "anglicised_terms_not_italic"
    End If

    If IsRuleEnabled(config, "foreign_names_not_italic") Then
        PerfTimerStart "foreign_names_not_italic"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Italics.Check_ForeignNamesNotItalic", doc)
        PerfTimerEnd "foreign_names_not_italic"
    End If

RunnerCleanup:
    ' -- Restore application state (always runs) ----------------
    On Error Resume Next
    Application.ScreenUpdating = wasScreenUpdating
    Application.StatusBar = wasStatusBar
    On Error GoTo 0

    ' -- Filter out issues inside block quotes / quoted text -----
    On Error Resume Next
    PerfTimerStart "FilterBlockQuoteIssues"
    Set allIssues = FilterBlockQuoteIssues(doc, allIssues)
    If Err.Number <> 0 Then
        ruleErrorCount = ruleErrorCount + 1
        ruleErrorLog = ruleErrorLog & "FilterBlockQuoteIssues (Err " & Err.Number & ": " & Err.Description & ")" & vbCrLf
        DebugLogError "RunAllPleadingsRules", "FilterBlockQuoteIssues", Err.Number, Err.Description
        Err.Clear
    End If
    PerfTimerEnd "FilterBlockQuoteIssues"
    On Error GoTo 0

    ' -- Print performance summary --------------------------------
    If ENABLE_PROFILING Then
        Dim perfSummary As String
        perfSummary = GetPerformanceSummary()
        Debug.Print perfSummary
    End If

    TraceStep "RunAllPleadingsRules", "total issues: " & allIssues.Count & _
              ", rule errors: " & ruleErrorCount
    TraceExit "RunAllPleadingsRules", allIssues.Count & " issues"

    Set RunAllPleadingsRules = allIssues
End Function

' ============================================================
'  FILTER: Remove issues inside block quotes, cover pages,
'  and contents/table-of-contents pages
'
'  Block quotes detected by:
'    1. Style name containing "quote", "block", or "extract"
'    2. Significant left indentation (> 36pt) with smaller font
'    3. Paragraph text wrapped in quotation marks
'
'  Cover pages detected by:
'    - Content before the first section break, OR
'    - All page-1 content when the document has > 1 page and
'      page 1 contains no numbered paragraphs
'
'  Contents pages detected by:
'    - Word's built-in TOC field ranges
'    - Paragraphs styled with "TOC" styles
'    - Paragraphs containing dot/tab leaders followed by numbers
' ============================================================
Private Function FilterBlockQuoteIssues(doc As Document, _
                                         issues As Collection) As Collection
    TraceEnter "FilterBlockQuoteIssues"
    TraceStep "FilterBlockQuoteIssues", "input: " & issues.Count & " issues"
    Dim filtered As New Collection
    Dim i As Long

    ' -- Determine cover page end position -------------------------
    ' Skip all content before the first "body text" paragraph,
    ' defined as the first paragraph whose plain text (without line
    ' breaks) exceeds BODY_TEXT_MIN_LEN characters.  Everything
    ' before that is treated as cover / title page.
    Const BODY_TEXT_MIN_LEN As Long = 200
    Dim coverPageEnd As Long
    coverPageEnd = -1  ' -1 means no cover page detected

    On Error Resume Next
    Dim coverPara As Paragraph
    For Each coverPara In doc.Paragraphs
        Err.Clear
        Dim cpText As String
        cpText = ""
        cpText = coverPara.Range.Text
        If Err.Number <> 0 Then Err.Clear: GoTo NextCoverPara
        ' Strip paragraph mark
        If Len(cpText) > 0 Then
            If Right$(cpText, 1) = vbCr Or Right$(cpText, 1) = Chr(13) Then
                cpText = Left$(cpText, Len(cpText) - 1)
            End If
        End If
        ' Strip any internal line breaks (vbLf, vertical tab, manual line break)
        Dim cleanCpText As String
        cleanCpText = Replace(Replace(Replace(cpText, vbLf, ""), vbVerticalTab, ""), Chr(11), "")
        If Len(cleanCpText) > BODY_TEXT_MIN_LEN Then
            ' This paragraph is the start of body text
            coverPageEnd = coverPara.Range.Start
            Exit For
        End If
NextCoverPara:
    Next coverPara
    On Error GoTo 0

    ' -- Determine TOC / contents page ranges -----------------------
    Dim tocStarts() As Long, tocEnds() As Long
    Dim tocCount As Long, tocCap As Long
    tocCap = 16
    ReDim tocStarts(0 To tocCap - 1)
    ReDim tocEnds(0 To tocCap - 1)
    tocCount = 0

    On Error Resume Next

    ' Method 1: Word's built-in TOC fields
    Dim toc As TableOfContents
    For Each toc In doc.TablesOfContents
        Err.Clear
        Dim tocRng As Range
        Set tocRng = toc.Range
        If Err.Number = 0 Then
            If tocCount >= tocCap Then
                tocCap = tocCap * 2
                ReDim Preserve tocStarts(0 To tocCap - 1)
                ReDim Preserve tocEnds(0 To tocCap - 1)
            End If
            tocStarts(tocCount) = tocRng.Start
            tocEnds(tocCount) = tocRng.End
            tocCount = tocCount + 1
        Else
            Err.Clear
        End If
    Next toc

    ' Method 2: Scan for TOC-styled paragraphs (catches manual TOCs)
    Dim tocPara As Paragraph
    For Each tocPara In doc.Paragraphs
        Err.Clear
        Dim tocSn As String
        tocSn = ""
        tocSn = LCase(tocPara.Style.NameLocal)
        If Err.Number <> 0 Then tocSn = "": Err.Clear

        Dim isTocPara As Boolean
        isTocPara = False

        ' Check style name for TOC indicators
        If InStr(tocSn, "toc") > 0 Or InStr(tocSn, "table of contents") > 0 Or _
           InStr(tocSn, "contents") > 0 Then
            isTocPara = True
        End If

        ' Check for dot/tab leader pattern: text followed by dots/tabs then page number
        If Not isTocPara Then
            Dim tocParaText As String
            tocParaText = ""
            tocParaText = tocPara.Range.Text
            If Err.Number <> 0 Then tocParaText = "": Err.Clear
            If Len(tocParaText) > 3 Then
                ' Pattern: dots or tabs followed by digits at end of line
                If tocParaText Like "*[." & vbTab & "][." & vbTab & "]*#" & vbCr Or _
                   tocParaText Like "*[." & vbTab & "][." & vbTab & "]*#" Then
                    isTocPara = True
                End If
            End If
        End If

        If isTocPara Then
            Dim tpStart As Long, tpEnd As Long
            tpStart = tocPara.Range.Start
            tpEnd = tocPara.Range.End
            If Err.Number = 0 Then
                If tocCount >= tocCap Then
                    tocCap = tocCap * 2
                    ReDim Preserve tocStarts(0 To tocCap - 1)
                    ReDim Preserve tocEnds(0 To tocCap - 1)
                End If
                tocStarts(tocCount) = tpStart
                tocEnds(tocCount) = tpEnd
                tocCount = tocCount + 1
            Else
                Err.Clear
            End If
        End If
    Next tocPara
    On Error GoTo 0

    ' -- Build list of block-quote paragraph ranges ----------------
    ' Detects block quotes via style name, indentation+smaller font,
    ' or multi-paragraph smart-quote spans (open " on first para,
    ' close " on last para — all paras in between are block-quoted).
    Dim bqStarts() As Long, bqEnds() As Long
    Dim bqCount As Long, bqCap As Long
    bqCap = 64
    ReDim bqStarts(0 To bqCap - 1)
    ReDim bqEnds(0 To bqCap - 1)
    bqCount = 0

    Dim insideMultiParaQuote As Boolean
    insideMultiParaQuote = False

    On Error Resume Next
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        Err.Clear
        Dim pStart As Long, pEnd As Long
        pStart = para.Range.Start
        pEnd = para.Range.End
        If Err.Number <> 0 Then Err.Clear: GoTo NxtBQ

        Dim isBQ As Boolean
        isBQ = False

        ' Check 1: Style name
        Dim sn As String
        sn = ""
        sn = LCase(para.Style.NameLocal)
        If Err.Number <> 0 Then sn = "": Err.Clear
        If InStr(sn, "quote") > 0 Or InStr(sn, "block") > 0 Or _
           InStr(sn, "extract") > 0 Then
            isBQ = True
        End If

        ' Check 1.5: Skip lists (mirrors IsBlockQuotePara CHECK 0)
        If Not isBQ Then
            Dim listLvl As Long
            listLvl = 0
            listLvl = para.Range.ListFormat.ListLevelNumber
            If Err.Number <> 0 Then listLvl = 0: Err.Clear
            If listLvl > 0 Then GoTo NxtBQ  ' Listed paragraph - not a block quote

            ' Check for bullet/number prefix in text
            Dim bqPText As String
            bqPText = ""
            bqPText = para.Range.Text
            If Err.Number <> 0 Then bqPText = "": Err.Clear
            If Len(bqPText) > 0 Then
                Dim fc As String
                fc = Left$(bqPText, 1)
                ' Bullet characters
                If fc = Chr(183) Or fc = ChrW(8226) Or fc = "-" Or fc = "*" Then GoTo NxtBQ
                ' Numbered list pattern: digit(s) followed by . or )
                If fc >= "0" And fc <= "9" Then
                    If bqPText Like "#[.)]#*" Or bqPText Like "##[.)]#*" Then GoTo NxtBQ
                End If
            End If
        End If

        ' Check 2: Indentation + smaller font or italic
        If Not isBQ Then
            Dim leftInd As Single
            leftInd = para.Format.LeftIndent
            If Err.Number <> 0 Then leftInd = 0: Err.Clear
            Dim fontSize As Single
            fontSize = para.Range.Font.Size
            If Err.Number <> 0 Then fontSize = 0: Err.Clear
            Dim bqItalic As Boolean
            bqItalic = False
            Dim bqItalVal As Long
            bqItalVal = para.Range.Font.Italic
            If Err.Number <> 0 Then bqItalVal = 0: Err.Clear
            If bqItalVal = -1 Then bqItalic = True  ' wdTrue = -1
            ' Moderate indent with clearly smaller font
            If leftInd > 18 And fontSize > 0 And fontSize < 11 Then
                isBQ = True
            End If
            ' Moderate indent with italic
            If leftInd > 18 And bqItalic Then
                isBQ = True
            End If
            ' Heavy indentation: only if italic or smaller font
            ' (plain indented body-size text = list, not quote)
            If Not isBQ And leftInd > 72 Then
                If bqItalic Or (fontSize > 0 And fontSize < 11) Then
                    isBQ = True
                End If
            End If
        End If

        ' Check 3: Multi-paragraph smart-quote detection
        Dim pText As String
        pText = ""
        pText = para.Range.Text
        If Err.Number <> 0 Then pText = "": Err.Clear
        ' Strip tabs, non-breaking spaces, CRs so quote marks are first/last
        pText = Replace(Replace(Replace(pText, vbCr, ""), vbTab, ""), ChrW(160), "")
        pText = Trim$(pText)
        If Not isBQ Then
            If Len(pText) > 2 Then
                Dim firstCh As Long, lastCh As Long
                Dim trimmed As String
                firstCh = AscW(Left(pText, 1))
                trimmed = pText
                If Right(trimmed, 1) = vbCr Or Right(trimmed, 1) = vbLf Then
                    trimmed = Left(trimmed, Len(trimmed) - 1)
                End If

                If Len(trimmed) > 1 Then
                    lastCh = AscW(Right(trimmed, 1))
                    ' Single-paragraph quote
                    If (firstCh = 8220 And lastCh = 8221) Then isBQ = True
                    If (firstCh = 34 And lastCh = 34) Then isBQ = True
                    ' Start of multi-paragraph quote (opens but doesn't close)
                    If Not isBQ And Not insideMultiParaQuote Then
                        If (firstCh = 8220 And lastCh <> 8221) Or _
                           (firstCh = 34 And lastCh <> 34) Then
                            insideMultiParaQuote = True
                            isBQ = True
                        End If
                    End If
                End If
            End If
        End If

        ' If inside a multi-paragraph quote, mark as block quote
        If insideMultiParaQuote And Not isBQ Then
            isBQ = True
        End If

        ' Check if this paragraph ends the multi-paragraph quote
        If insideMultiParaQuote And Len(pText) > 1 Then
            Dim endTrimmed As String
            endTrimmed = pText
            If Right(endTrimmed, 1) = vbCr Or Right(endTrimmed, 1) = vbLf Then
                endTrimmed = Left(endTrimmed, Len(endTrimmed) - 1)
            End If
            If Len(endTrimmed) > 0 Then
                Dim endCh As Long
                endCh = AscW(Right(endTrimmed, 1))
                If endCh = 8221 Or endCh = 34 Then
                    insideMultiParaQuote = False
                End If
            End If
        End If

        If isBQ Then
            If bqCount >= bqCap Then
                bqCap = bqCap * 2
                ReDim Preserve bqStarts(0 To bqCap - 1)
                ReDim Preserve bqEnds(0 To bqCap - 1)
            End If
            bqStarts(bqCount) = pStart
            bqEnds(bqCount) = pEnd
            bqCount = bqCount + 1
        End If
NxtBQ:
    Next para
    On Error GoTo 0

    ' -- Filter issues ---------------------------------------------
    If bqCount = 0 And coverPageEnd < 0 And tocCount = 0 Then
        Set FilterBlockQuoteIssues = issues
        Exit Function
    End If

    For i = 1 To issues.Count
        Dim finding As Object
        Set finding = issues(i)
        Dim rs As Long
        rs = GetIssueProp(finding, "RangeStart")

        ' Skip issues on cover page
        If coverPageEnd > 0 And rs < coverPageEnd Then GoTo SkipIssue

        ' Skip issues in table of contents / contents pages
        Dim inTOC As Boolean
        inTOC = False
        Dim t As Long
        For t = 0 To tocCount - 1
            If rs >= tocStarts(t) And rs < tocEnds(t) Then
                inTOC = True
                Exit For
            End If
        Next t
        If inTOC Then GoTo SkipIssue

        ' Skip content-based issues in block quotes
        ' (formatting rules like font_consistency still apply)
        Dim inBQ As Boolean
        inBQ = False
        Dim j As Long
        For j = 0 To bqCount - 1
            If rs >= bqStarts(j) And rs < bqEnds(j) Then
                inBQ = True
                Exit For
            End If
        Next j
        ' Suppress ALL rules in block quotes
        If inBQ Then GoTo SkipIssue

        filtered.Add finding
        GoTo NextIssue
SkipIssue:
NextIssue:
    Next i

    TraceStep "FilterBlockQuoteIssues", "output: " & filtered.Count & " issues (" & _
              (issues.Count - filtered.Count) & " filtered out)"
    TraceExit "FilterBlockQuoteIssues"
    Set FilterBlockQuoteIssues = filtered
End Function

' ============================================================
'  APPLY HIGHLIGHTS AND COMMENTS
' ============================================================
Public Sub ApplyHighlights(doc As Document, _
                           issues As Collection, _
                           Optional addComments As Boolean = True)
    TraceEnter "ApplyHighlights"
    DebugLogDoc "ApplyHighlights target", doc
    TraceStep "ApplyHighlights", issues.Count & " issues, addComments=" & addComments

    Dim finding As Object
    Dim rng As Range
    Dim i As Long
    Dim cmtRef As Comment

    ' Suppress screen updates during batch comment insertion
    Dim wasScreenUpdating As Boolean
    wasScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim wasStatusBar As Variant
    wasStatusBar = Application.StatusBar

    On Error GoTo HighlightCleanup

    For i = 1 To issues.Count
        Set finding = issues(i)
        If GetIssueProp(finding, "RangeStart") >= 0 And GetIssueProp(finding, "RangeEnd") > GetIssueProp(finding, "RangeStart") Then
            On Error Resume Next: Err.Clear
            Set rng = doc.Range(GetIssueProp(finding, "RangeStart"), GetIssueProp(finding, "RangeEnd"))
            If Err.Number = 0 Then
                ' Apply yellow highlight to the flagged range
                rng.HighlightColorIndex = wdYellow
                If Err.Number <> 0 Then
                    DebugLogError "ApplyHighlights", "highlight i=" & i, Err.Number, Err.Description
                    Err.Clear
                End If
                If addComments Then
                    TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                        "ApplyHighlights", "comment i=" & i
                End If
            Else
                DebugLogError "ApplyHighlights", "doc.Range i=" & i & _
                    " start=" & GetIssueProp(finding, "RangeStart") & _
                    " end=" & GetIssueProp(finding, "RangeEnd"), Err.Number, Err.Description
                Err.Clear
            End If
            On Error GoTo HighlightCleanup
        Else
            TraceStep "ApplyHighlights", "SKIPPED i=" & i & _
                      " -- invalid range start=" & GetIssueProp(finding, "RangeStart") & _
                      " end=" & GetIssueProp(finding, "RangeEnd")
        End If
    Next i

HighlightCleanup:
    On Error Resume Next
    Application.ScreenUpdating = wasScreenUpdating
    Application.StatusBar = wasStatusBar
    On Error GoTo 0
    TraceExit "ApplyHighlights"
End Sub

' ============================================================
'  APPLY SUGGESTIONS VIA TRACKED CHANGES
' ============================================================
Public Sub ApplySuggestionsAsTrackedChanges(doc As Document, _
                                             issues As Collection, _
                                             Optional addComments As Boolean = True)
    TraceEnter "ApplyTrackedChanges"
    DebugLogDoc "ApplyTrackedChanges target", doc
    TraceStep "ApplyTrackedChanges", issues.Count & " issues, addComments=" & addComments

    Dim finding As Object
    Dim rng As Range
    Dim i As Long
    Dim cmtRef As Comment
    Dim wasTrackingChanges As Boolean
    wasTrackingChanges = doc.TrackRevisions

    ' Suppress screen updates during batch application to prevent
    ' Word from repaginating/redrawing after each comment/tracked change
    Dim wasScreenUpdating As Boolean
    wasScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' Capture status bar so we can restore it in cleanup
    Dim wasStatusBar As Variant
    wasStatusBar = Application.StatusBar

    ' Enable tracking for the entire batch; restored once in cleanup.
    doc.TrackRevisions = True

    On Error GoTo TrackedCleanup

    ' Process from end of document backwards so tracked-change
    ' insertions / deletions do not shift positions of later issues
    For i = issues.Count To 1 Step -1
        Set finding = issues(i)
        If GetIssueProp(finding, "RangeStart") >= 0 And GetIssueProp(finding, "RangeEnd") > GetIssueProp(finding, "RangeStart") Then
            On Error Resume Next: Err.Clear
            Set rng = doc.Range(GetIssueProp(finding, "RangeStart"), GetIssueProp(finding, "RangeEnd"))
            If Err.Number = 0 Then
                If GetIssueProp(finding, "AutoFixSafe") Then
                    ' Remember original position and length before modification
                    Dim origStart As Long
                    Dim origLen As Long
                    Dim sugText As String
                    origStart = rng.Start
                    origLen = rng.End - rng.Start
                    ' Prefer ReplacementText (literal replacement) over Suggestion (human-readable)
                    sugText = ""
                    sugText = CStr(GetIssueProp(finding, "ReplacementText"))
                    If Len(sugText) = 0 Then sugText = GetIssueProp(finding, "Suggestion")

                    ' --- WHITESPACE VALIDATION GATE ---
                    Dim origText As String
                    origText = rng.Text
                    If Err.Number <> 0 Then origText = "": Err.Clear

                    Dim skipAmendment As Boolean
                    skipAmendment = False

                    ' For deletions (empty suggestion = delete the range)
                    If Len(sugText) = 0 And Len(origText) > 0 Then
                        Dim chIdx As Long
                        Dim ch As String
                        For chIdx = 1 To Len(origText)
                            ch = Mid$(origText, chIdx, 1)
                            If (ch >= "A" And ch <= "Z") Or _
                               (ch >= "a" And ch <= "z") Or _
                               (ch >= "0" And ch <= "9") Or _
                               ch = "." Then
                                skipAmendment = True
                                Debug.Print "WHITESPACE VALIDATION: Skipped deletion of '" & origText & "' -- contains substantive character '" & ch & "'"
                                Exit For
                            End If
                        Next chIdx
                    End If

                    ' For replacements, verify we are only changing whitespace
                    If Len(sugText) > 0 And Len(origText) > 0 Then
                        Dim isOnlyWhitespace As Boolean
                        isOnlyWhitespace = True
                        For chIdx = 1 To Len(origText)
                            ch = Mid$(origText, chIdx, 1)
                            If ch <> " " And ch <> vbTab And ch <> ChrW(160) Then
                                isOnlyWhitespace = False
                                Exit For
                            End If
                        Next chIdx

                        If Not isOnlyWhitespace Then
                            If Len(sugText) < Len(origText) Then
                                Dim origHasPeriod As Boolean
                                origHasPeriod = (InStr(1, origText, ".") > 0)
                                Dim sugHasPeriod As Boolean
                                sugHasPeriod = (InStr(1, sugText, ".") > 0)
                                If origHasPeriod And Not sugHasPeriod Then
                                    skipAmendment = True
                                    Debug.Print "WHITESPACE VALIDATION: Skipped replacement '" & origText & "' -> '" & sugText & "' -- would remove period"
                                End If
                            End If
                        End If
                    End If

                    If skipAmendment Then
                        TraceStep "ApplyTrackedChanges", "SKIPPED amendment i=" & i & _
                                  " orig=""" & Left$(origText, 30) & """ sug=""" & Left$(sugText, 30) & """"
                        If addComments Then
                            TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                                "ApplyTrackedChanges", "skip-comment i=" & i
                        End If
                        GoTo NextApplyIssue
                    End If

                    ' Apply tracked change
                    TraceStep "ApplyTrackedChanges", "APPLYING i=" & i & _
                              " range=" & origStart & "-" & (origStart + origLen) & _
                              " orig=""" & Left$(origText, 30) & """ -> """ & Left$(sugText, 30) & """"
                    TrySetRangeText rng, sugText, _
                        "ApplyTrackedChanges", "apply i=" & i
                Else
                    If addComments Then
                        TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                            "ApplyTrackedChanges", "comment-only i=" & i
                    End If
                End If
            Else
                DebugLogError "ApplyTrackedChanges", "doc.Range i=" & i & _
                    " start=" & GetIssueProp(finding, "RangeStart") & _
                    " end=" & GetIssueProp(finding, "RangeEnd"), Err.Number, Err.Description
                Err.Clear
            End If
NextApplyIssue:
            On Error GoTo TrackedCleanup
        Else
            TraceStep "ApplyTrackedChanges", "SKIPPED i=" & i & _
                      " -- invalid range start=" & GetIssueProp(finding, "RangeStart") & _
                      " end=" & GetIssueProp(finding, "RangeEnd")
        End If
    Next i

TrackedCleanup:
    ' Single cleanup path: always restore document and application state.
    On Error Resume Next
    doc.TrackRevisions = wasTrackingChanges
    Application.ScreenUpdating = wasScreenUpdating
    Application.StatusBar = wasStatusBar
    On Error GoTo 0
    TraceExit "ApplyTrackedChanges"
End Sub

' ============================================================
'  PRIVATE: Build comment text from an issue dictionary
' ============================================================
Private Function BuildCommentText(finding As Object) As String
    Dim txt As String
    txt = GetIssueProp(finding, "Issue")
    Dim sug As String
    sug = GetIssueProp(finding, "Suggestion")
    ' Only append suggestion text if it's human-readable (not a literal replacement)
    If Len(sug) > 0 And Len(Trim(sug)) > 1 Then
        txt = txt & " -- Suggestion: " & sug
    End If
    BuildCommentText = txt
End Function

' ============================================================
'  GENERATE JSON REPORT
' ============================================================
Public Function GenerateReport(issues As Collection, _
                                filePath As String, _
                                Optional doc As Document = Nothing) As String
    TraceEnter "GenerateReport"
    TraceStep "GenerateReport", issues.Count & " issues, path=" & filePath

    Dim fileNum As Integer
    Dim finding As Object
    Dim i As Long

    ' Resolve document name: prefer explicit doc, fall back to ActiveDocument
    Dim docName As String
    On Error Resume Next
    If Not doc Is Nothing Then
        docName = doc.Name
    Else
        docName = ActiveDocument.Name
    End If
    If Err.Number <> 0 Then docName = "(unknown)": Err.Clear
    On Error GoTo 0

    ' Open file with error handling
    fileNum = FreeFile
    On Error Resume Next
    Open filePath For Output As #fileNum
    If Err.Number <> 0 Then
        GenerateReport = "Error: could not write to " & filePath & _
                         " (Err " & Err.Number & ": " & Err.Description & ")"
        DebugLogError "GenerateReport", "open " & filePath, Err.Number, Err.Description
        Err.Clear
        On Error GoTo 0
        TraceExit "GenerateReport", "FAILED open"
        Exit Function
    End If
    On Error GoTo 0

    On Error GoTo ReportWriteErr

    Print #fileNum, "{"
    Print #fileNum, "  ""document"": """ & EscJSON(docName) & ""","
    Print #fileNum, "  ""timestamp"": """ & Format(Now, "yyyy-mm-ddThh:nn:ss") & ""","
    Print #fileNum, "  ""total_issues"": " & issues.Count & ","

    Print #fileNum, "  ""issues"": ["
    For i = 1 To issues.Count
        Set finding = issues(i)
        If i < issues.Count Then
            Print #fileNum, IssueToJSON(finding) & ","
        Else
            Print #fileNum, IssueToJSON(finding)
        End If
    Next i
    Print #fileNum, "  ],"

    Dim countDict As Object
    Set countDict = CreateObject("Scripting.Dictionary")
    For i = 1 To issues.Count
        Set finding = issues(i)
        If countDict.Exists(GetIssueProp(finding, "RuleName")) Then
            countDict(GetIssueProp(finding, "RuleName")) = countDict(GetIssueProp(finding, "RuleName")) + 1
        Else
            countDict.Add GetIssueProp(finding, "RuleName"), 1
        End If
    Next i

    Print #fileNum, "  ""summary"": {"
    Print #fileNum, "    ""counts_per_rule"": {"
    Dim keys As Variant
    keys = countDict.keys
    Dim k As Long
    For k = 0 To countDict.Count - 1
        If k < countDict.Count - 1 Then
            Print #fileNum, "      """ & EscJSON(CStr(keys(k))) & """: " & countDict(keys(k)) & ","
        Else
            Print #fileNum, "      """ & EscJSON(CStr(keys(k))) & """: " & countDict(keys(k))
        End If
    Next k
    Print #fileNum, "    }"
    Print #fileNum, "  }"
    Print #fileNum, "}"

    Close #fileNum

    Dim summaryStr As String
    summaryStr = "Report saved: " & filePath & vbCrLf
    summaryStr = summaryStr & "Total issues: " & issues.Count
    GenerateReport = summaryStr
    TraceExit "GenerateReport", issues.Count & " issues written"
    Exit Function

ReportWriteErr:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    GenerateReport = "Error writing report: Err " & Err.Number & ": " & Err.Description
    DebugLogError "GenerateReport", "write", Err.Number, Err.Description
    TraceExit "GenerateReport", "FAILED"
End Function

' ============================================================
'  HUMAN-READABLE ISSUE SUMMARY
' ============================================================
Public Function GetIssueSummary(issues As Collection) As String
    Dim countDict As Object
    Set countDict = CreateObject("Scripting.Dictionary")
    Dim finding As Object
    Dim i As Long

    For i = 1 To issues.Count
        Set finding = issues(i)
        If countDict.Exists(GetIssueProp(finding, "RuleName")) Then
            countDict(GetIssueProp(finding, "RuleName")) = countDict(GetIssueProp(finding, "RuleName")) + 1
        Else
            countDict.Add GetIssueProp(finding, "RuleName"), 1
        End If
    Next i

    Dim result As String
    Dim keys As Variant
    Dim k As Long

    If countDict.Count = 0 Then
        GetIssueSummary = "No issues found."
        Exit Function
    End If

    keys = countDict.keys
    For k = 0 To countDict.Count - 1
        Dim cnt As Long
        cnt = countDict(keys(k))
        result = result & CStr(keys(k)) & ": " & cnt & " finding"
        If cnt <> 1 Then result = result & "s"
        result = result & vbCrLf
    Next k

    result = result & vbCrLf & "Total: " & issues.Count & " finding"
    If issues.Count <> 1 Then result = result & "s"
    GetIssueSummary = result
End Function

' ============================================================
'  RULE DISPLAY NAMES (for launcher summary)
' ============================================================
Public Function GetRuleDisplayNames() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    d.Add "spelling", "Spelling Enforcement (UK/US)"
    d.Add "repeated_words", "Repeated Word Detection"
    d.Add "sequential_numbering", "Sequential Numbering"
    d.Add "heading_capitalisation", "Heading Capitalisation"
    d.Add "custom_term_whitelist", "Custom Term Whitelist"
    d.Add "defined_terms", "Defined Term Checker"
    d.Add "clause_number_format", "Clause Number Format"
    d.Add "date_time_format", "Date/Time Format Consistency"
    d.Add "list_rules", "List Format & Punctuation"
    d.Add "formatting_consistency", "Formatting Consistency"
    d.Add "licence_license", "Licence/License Rule"
    d.Add "check_cheque", "Check/Cheque Rule"
    d.Add "slash_style", "Slash Style Checker"
    d.Add "dash_usage", "En-dash/Em-dash/Hyphen"
    d.Add "bracket_integrity", "Bracket Integrity"
    d.Add "quotation_mark_consistency", "Quotation Mark Consistency"
    d.Add "currency_number_format", "Currency/Number Formatting"
    d.Add "footnote_rules", "Footnote Rules"
    d.Add "title_formatting", "Title Formatting Consistency"
    d.Add "brand_name_enforcement", "Brand Name Enforcement"
    d.Add "mandated_legal_term_forms", "Mandated Legal Term Forms"
    d.Add "always_capitalise_terms", "Always Capitalise Terms"
    d.Add "known_anglicised_terms_not_italic", "Anglicised Terms Not Italic"
    d.Add "foreign_names_not_italic", "Foreign Names Not Italic"
    d.Add "single_quotes_default", "Single Quotes Default"
    d.Add "smart_quote_consistency", "Smart Quote Consistency"
    d.Add "spell_out_under_ten", "Spell Out Numbers Under 10"
    d.Add "double_spaces", "Double Spaces"
    d.Add "double_commas", "Double Commas"
    d.Add "space_before_punct", "Space Before Punctuation"
    d.Add "missing_space_after_dot", "Missing Space After Full Stop"
    d.Add "trailing_spaces", "Trailing Spaces"

    Set GetRuleDisplayNames = d
End Function

' ============================================================
'  CONFIG DRIFT VALIDATION (development helper)
'  Call from Immediate window: PleadingsEngine.ValidateConfigDrift
'  Prints any keys present in config but missing from display
'  names, or vice versa.
' ============================================================
Public Sub ValidateConfigDrift()
    Dim cfg As Object
    Set cfg = InitRuleConfig()
    Dim disp As Object
    Set disp = GetRuleDisplayNames()
    Dim k As Variant
    Dim driftFound As Boolean
    driftFound = False

    For Each k In cfg.keys
        If Not disp.Exists(CStr(k)) Then
            Debug.Print "DRIFT: config key '" & k & "' has no display name"
            driftFound = True
        End If
    Next k

    For Each k In disp.keys
        If Not cfg.Exists(CStr(k)) Then
            Debug.Print "DRIFT: display name '" & k & "' has no config key"
            driftFound = True
        End If
    Next k

    If Not driftFound Then
        Debug.Print "ValidateConfigDrift: OK -- config and display names are in sync"
    End If
End Sub

' ============================================================
'  HELPERS: PAGE RANGE
'  Accepts flexible page specifications:
'    "5"         - single page
'    "3-7"       - range (also supports en-dash and colon)
'    "1,3,5"     - comma-separated pages
'    "1,3-5,8"   - mixed
'    ""          - all pages (no filter)
' ============================================================
Public Function IsInPageRange(rng As Range) As Boolean
    If pageRangeSet Is Nothing Then
        IsInPageRange = True
        Exit Function
    End If
    If pageRangeSet.Count = 0 Then
        IsInPageRange = True
        Exit Function
    End If
    Dim pageNum As Long
    pageNum = rng.Information(wdActiveEndAdjustedPageNumber)
    IsInPageRange = pageRangeSet.Exists(pageNum)
End Function

Public Sub SetPageRange(startPage As Long, endPage As Long)
    ' Legacy compatibility: convert start/end to page set
    If startPage = 0 And endPage = 0 Then
        Set pageRangeSet = Nothing
        Exit Sub
    End If
    Set pageRangeSet = CreateObject("Scripting.Dictionary")
    Dim pg As Long
    For pg = startPage To endPage
        pageRangeSet(pg) = True
    Next pg
End Sub

Public Sub SetPageRangeFromString(ByVal spec As String)
    ' Parse flexible page range specification
    spec = Trim(spec)
    If Len(spec) = 0 Then
        Set pageRangeSet = Nothing
        Exit Sub
    End If

    Set pageRangeSet = CreateObject("Scripting.Dictionary")

    ' Normalise separators: en-dash (8211) and colon to hyphen
    spec = Replace(spec, ChrW(8211), "-")
    spec = Replace(spec, ":", "-")

    ' Split on comma
    Dim parts() As String
    parts = Split(spec, ",")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim part As String
        part = Trim(parts(i))
        If Len(part) = 0 Then GoTo NextPart

        Dim dashPos As Long
        dashPos = InStr(1, part, "-")

        If dashPos > 0 Then
            ' Range: "3-7"
            Dim rangeStart As Long
            Dim rangeEnd As Long
            Dim leftPart As String
            Dim rightPart As String
            leftPart = Trim(Left$(part, dashPos - 1))
            rightPart = Trim(Mid$(part, dashPos + 1))

            If IsNumeric(leftPart) And IsNumeric(rightPart) Then
                rangeStart = CLng(leftPart)
                rangeEnd = CLng(rightPart)
                Dim pg As Long
                For pg = rangeStart To rangeEnd
                    pageRangeSet(pg) = True
                Next pg
            End If
        Else
            ' Single page: "5"
            If IsNumeric(part) Then
                pageRangeSet(CLng(part)) = True
            End If
        End If
NextPart:
    Next i

    ' If nothing valid was parsed, clear the set
    If pageRangeSet.Count = 0 Then
        Set pageRangeSet = Nothing
    End If
End Sub

Public Function GetRuleErrorCount() As Long
    GetRuleErrorCount = ruleErrorCount
End Function

Public Function GetRuleErrorLog() As String
    GetRuleErrorLog = ruleErrorLog
End Function

' ============================================================
'  HELPERS: WHITELIST
' ============================================================
Public Function IsWhitelistedTerm(term As String) As Boolean
    If whitelistDict Is Nothing Then
        IsWhitelistedTerm = False
        Exit Function
    End If
    IsWhitelistedTerm = whitelistDict.Exists(LCase(term))
End Function

Public Sub SetWhitelist(terms As Object)
    Set whitelistDict = terms
End Sub

' ============================================================
'  HELPERS: LOCATION STRING
' ============================================================
Public Function GetLocationString(rng As Range, doc As Document) As String
    Dim pageNum As Long
    Dim paraNum As Long

    On Error Resume Next
    pageNum = rng.Information(wdActiveEndAdjustedPageNumber)
    If Err.Number <> 0 Then pageNum = 0: Err.Clear
    On Error GoTo 0

    ' Use cached paragraph positions for O(log N) lookup
    ' instead of iterating all paragraphs (O(N) per call)
    paraNum = FindParagraphIndex(rng.Start)

    GetLocationString = "page " & pageNum & " paragraph " & paraNum
End Function

' ============================================================
'  PRIVATE HELPERS
' ============================================================
Private Function IsRuleEnabled(config As Object, _
                                ruleName As String) As Boolean
    If config.Exists(ruleName) Then
        IsRuleEnabled = CBool(config(ruleName))
    Else
        IsRuleEnabled = False
    End If
End Function

Private Sub AddIssuesToCollection(master As Collection, _
                                   ruleIssues As Collection)
    Dim i As Long
    If ruleIssues Is Nothing Then Exit Sub
    For i = 1 To ruleIssues.Count
        master.Add ruleIssues(i)
    Next i
End Sub

Private Function EscJSON(ByVal txt As String) As String
    txt = Replace(txt, "\", "\\")
    txt = Replace(txt, """", "\""")
    txt = Replace(txt, vbCr, "\r")
    txt = Replace(txt, vbLf, "\n")
    txt = Replace(txt, vbTab, "\t")
    EscJSON = txt
End Function

' ================================================================
'  PUBLIC: Factory function to create a dictionary-based finding
'  Called by rule modules via Application.Run
' ================================================================
Public Function CreateIssue(ByVal ruleName_ As String, _
                            ByVal location_ As String, _
                            ByVal issue_ As String, _
                            ByVal suggestion_ As String, _
                            ByVal rangeStart_ As Long, _
                            ByVal rangeEnd_ As Long, _
                            Optional ByVal severity_ As String = "error", _
                            Optional ByVal autoFixSafe_ As Boolean = False, _
                            Optional ByVal replacementText_ As String = "") As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    d("ReplacementText") = replacementText_
    Set CreateIssue = d
End Function

' ================================================================
'  PRIVATE: Read a property from an finding (supports both
'  issue dictionary class and Dictionary-based issues)
' ================================================================
Private Function GetIssueProp(finding As Object, ByVal propName As String) As Variant
    On Error Resume Next
    ' Try dictionary access first
    If TypeName(finding) = "Dictionary" Then
        GetIssueProp = finding(propName)
    Else
        ' Fall back to object property access via CallByName
        GetIssueProp = CallByName(finding, propName, VbGet)
    End If
    If Err.Number <> 0 Then
        GetIssueProp = ""
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ================================================================
'  PRIVATE: Format an finding as JSON (supports both types)
' ================================================================
Private Function IssueToJSON(finding As Object) As String
    Dim s As String
    s = "    {" & vbCrLf
    s = s & "      ""rule"": """ & EscJSON(CStr(GetIssueProp(finding, "RuleName"))) & """," & vbCrLf
    s = s & "      ""location"": """ & EscJSON(CStr(GetIssueProp(finding, "Location"))) & """," & vbCrLf
    s = s & "      ""severity"": """ & EscJSON(CStr(GetIssueProp(finding, "Severity"))) & """," & vbCrLf
    s = s & "      ""finding"": """ & EscJSON(CStr(GetIssueProp(finding, "Issue"))) & """," & vbCrLf
    s = s & "      ""suggestion"": """ & EscJSON(CStr(GetIssueProp(finding, "Suggestion"))) & """," & vbCrLf
    Dim repText As String
    repText = CStr(GetIssueProp(finding, "ReplacementText"))
    If Len(repText) > 0 Then
        s = s & "      ""replacement_text"": """ & EscJSON(repText) & """," & vbCrLf
    End If
    s = s & "      ""auto_fix_safe"": " & IIf(CBool(GetIssueProp(finding, "AutoFixSafe")), "true", "false") & vbCrLf
    s = s & "    }"
    IssueToJSON = s
End Function

```

# FILE: PleadingsLauncher.bas

```vb
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
    Set issues = Application.Run("PleadingsEngine.RunAllPleadingsRules", ActiveDocument, cfg)

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

    Dim applyChoice As Long
    applyChoice = MsgBox(issues.Count & " issue(s) found." & errInfo & vbCrLf & vbCrLf & _
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
            brandDir = modDebugLog.GetParentDirectory(savePath)
            If Len(brandDir) > 0 Then
                modDebugLog.EnsureDirectoryExists brandDir
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
    If ActiveDocument.Path <> "" Then
        Dim baseName As String
        baseName = ActiveDocument.Name
        Dim dotPos As Long
        dotPos = InStrRev(baseName, ".")
        If dotPos > 1 Then baseName = Left$(baseName, dotPos - 1)
        reportPath = ActiveDocument.Path & sep & baseName & "_pleadings_report.json"
    End If
    If Err.Number <> 0 Or Len(reportPath) = 0 Then
        Err.Clear
        reportPath = ""
    End If
    On Error GoTo 0

    If Len(reportPath) = 0 Then
        Dim tmpDir As String
        #If Mac Then
            tmpDir = Environ("TMPDIR")
            If Len(tmpDir) = 0 Then tmpDir = "/tmp"
            If Right$(tmpDir, 1) = sep Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
        #Else
            tmpDir = Environ("TEMP")
            If Len(tmpDir) = 0 Then tmpDir = Environ("TMP")
            If Len(tmpDir) = 0 Then tmpDir = Environ("USERPROFILE")
            If Len(tmpDir) = 0 Then tmpDir = "C:\Temp"
            If Right$(tmpDir, 1) = sep Then tmpDir = Left$(tmpDir, Len(tmpDir) - 1)
        #End If
        reportPath = tmpDir & sep & "pleadings_report.json"
    End If

    ' Ensure parent directory exists before writing
    Dim reportDir As String
    reportDir = modDebugLog.GetParentDirectory(reportPath)
    If Len(reportDir) > 0 Then
        modDebugLog.EnsureDirectoryExists reportDir
    End If

    Dim summary As String
    summary = Application.Run("PleadingsEngine.GenerateReport", issues, reportPath, ActiveDocument)

    ' Auto-save debug log alongside report when DEBUG_MODE is True
    Dim logPath As String
    Dim logSaved As Boolean
    logSaved = False
    logPath = ""

    On Error Resume Next
    If modDebugLog.DEBUG_MODE Then
        logPath = Left$(reportPath, Len(reportPath) - 5) & "_debug.log"
        logSaved = modDebugLog.DebugLogSaveToTextFile(logPath)
    End If
    On Error GoTo 0

    Dim msg As String
    msg = "Report saved to:" & vbCrLf & reportPath

    If logSaved And Len(logPath) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Debug log saved to:" & vbCrLf & logPath
    ElseIf modDebugLog.DEBUG_MODE And Not logSaved Then
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


```

# FILE: Rules_Brands.bas

```vb
Attribute VB_Name = "Rules_Brands"
' ============================================================
' Rules_Brands.bas
' Combined module for Rule 22: Brand Name Enforcement
'
' Proofreading rule: enforces correct brand/entity name
' spellings and capitalisations. Maintains a configurable
' dictionary of correct forms and their known incorrect
' brandVariants, and flags any incorrect usage found in the
' document.
'
' Provides persistence via SaveBrandRules / LoadBrandRules
' for user-customised brand lists.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "brand_name_enforcement"

' -- Module-level brand rules dictionary ---------------------
' Key = correct form (String), Value = comma-separated incorrect brandVariants (String)
Private brandRules As Object

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Function Check_BrandNameEnforcement(doc As Document) As Collection
    Dim issues As New Collection

    ' Initialise defaults if not yet loaded
    If brandRules Is Nothing Then
        InitDefaultBrands
    End If

    Dim keys As Variant
    Dim k As Long
    Dim correctForm As String
    Dim brandVariants As Variant
    Dim v As Long
    Dim brandVariant As String

    keys = brandRules.keys

    For k = 0 To brandRules.Count - 1
        correctForm = CStr(keys(k))
        brandVariants = Split(CStr(brandRules(correctForm)), ",")

        For v = LBound(brandVariants) To UBound(brandVariants)
            brandVariant = Trim(CStr(brandVariants(v)))
            If Len(brandVariant) = 0 Then GoTo NextVariant

            SearchAndFlag doc, brandVariant, correctForm, issues

NextVariant:
        Next v
    Next k

    Set Check_BrandNameEnforcement = issues
End Function

' ============================================================
'  PRIVATE: Search for an incorrect brandVariant and flag matches
' ============================================================
Private Sub SearchAndFlag(doc As Document, _
                           brandVariant As String, _
                           correctForm As String, _
                           ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = brandVariant
        .MatchWholeWord = True
        .MatchCase = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        If EngineIsInPageRange(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME, locStr, "Incorrect brand name: '" & rng.Text & "'", "Use '" & correctForm & "'", rng.Start, rng.End, "error")
            issues.Add finding
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ============================================================
'  PRIVATE: Populate default brand rules
' ============================================================
Private Sub InitDefaultBrands()
    Set brandRules = CreateObject("Scripting.Dictionary")

    brandRules.Add "PwC", "PWC,Pwc,pwc"
    brandRules.Add "Deloitte", "deloitte,DELOITTE"
    brandRules.Add "HMRC", "Hmrc,hmrc,H.M.R.C."
    brandRules.Add "FCA", "Fca,fca,F.C.A."
    brandRules.Add "EY", "ey,Ernst & Young,Ernst and Young"
    brandRules.Add "KPMG", "kpmg,Kpmg"
End Sub

' ============================================================
'  PUBLIC: Add or update a brand rule
' ============================================================
Public Sub AddBrandRule(correct As String, incorrectVariants As String)
    If brandRules Is Nothing Then
        InitDefaultBrands
    End If

    If brandRules.Exists(correct) Then
        brandRules(correct) = incorrectVariants
    Else
        brandRules.Add correct, incorrectVariants
    End If
End Sub

' ============================================================
'  PUBLIC: Remove a brand rule
' ============================================================
Public Sub RemoveBrandRule(correct As String)
    If brandRules Is Nothing Then Exit Sub

    If brandRules.Exists(correct) Then
        brandRules.Remove correct
    End If
End Sub

' ============================================================
'  PUBLIC: Get current brand rules dictionary
' ============================================================
Public Function GetBrandRules() As Object
    If brandRules Is Nothing Then
        InitDefaultBrands
    End If

    Set GetBrandRules = brandRules
End Function

' ============================================================
'  PUBLIC: Save brand rules to a text file
'  Format: one line per rule -- "CorrectForm=brandVariant1,brandVariant2"
' ============================================================
Public Function SaveBrandRules(filePath As String) As Boolean
    SaveBrandRules = False
    If brandRules Is Nothing Then Exit Function

    Dim fileNum As Integer
    Dim keys As Variant
    Dim k As Long

    fileNum = FreeFile
    On Error GoTo SaveError
    Open filePath For Output As #fileNum

    keys = brandRules.keys
    For k = 0 To brandRules.Count - 1
        Print #fileNum, CStr(keys(k)) & "=" & CStr(brandRules(keys(k)))
    Next k

    Close #fileNum
    SaveBrandRules = True
    Exit Function

SaveError:
    On Error Resume Next
    Close #fileNum
    Debug.Print "Rules_Brands.SaveBrandRules: Err " & Err.Number & ": " & Err.Description
    On Error GoTo 0
End Function

' ============================================================
'  PUBLIC: Load brand rules from a text file
'  Replaces existing rules with contents of the file.
'  Format: one line per rule -- "CorrectForm=brandVariant1,brandVariant2"
' ============================================================
Public Function LoadBrandRules(filePath As String) As Boolean
    LoadBrandRules = False
    Dim fileNum As Integer
    Dim lineText As String
    Dim eqPos As Long
    Dim correct As String
    Dim brandVariants As String

    Set brandRules = CreateObject("Scripting.Dictionary")

    fileNum = FreeFile
    On Error GoTo LoadError
    Open filePath For Input As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim(lineText)

        ' Skip empty lines and comments
        If Len(lineText) = 0 Then GoTo NextLine
        If Left(lineText, 1) = "#" Then GoTo NextLine

        eqPos = InStr(lineText, "=")
        If eqPos > 1 Then
            correct = Trim(Left(lineText, eqPos - 1))
            brandVariants = Trim(Mid(lineText, eqPos + 1))

            If Len(correct) > 0 And Len(brandVariants) > 0 Then
                If Not brandRules.Exists(correct) Then
                    brandRules.Add correct, brandVariants
                End If
            End If
        End If

NextLine:
    Loop

    Close #fileNum
    LoadBrandRules = True
    Exit Function

LoadError:
    On Error Resume Next
    Close #fileNum
    Debug.Print "Rules_Brands.LoadBrandRules: Err " & Err.Number & ": " & Err.Description
    On Error GoTo 0
    ' If file could not be loaded, fall back to defaults.
    If brandRules Is Nothing Then
        InitDefaultBrands
    ElseIf brandRules.Count = 0 Then
        InitDefaultBrands
    End If
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ============================================================
'  PUBLIC: Default brand-rules file path (cross-platform)
'  Called by frmPleadingsChecker and PleadingsLauncher via
'  Application.Run so the path is defined in one place.
' ============================================================
Public Function GetDefaultBrandRulesPath() As String
    Dim sep As String
    sep = Application.PathSeparator
    #If Mac Then
        GetDefaultBrandRulesPath = Environ("HOME") & sep & "Library" & sep & _
                                    "Application Support" & sep & "PleadingsChecker" & sep & "brand_rules.txt"
    #Else
        GetDefaultBrandRulesPath = Environ("APPDATA") & sep & "PleadingsChecker" & sep & "brand_rules.txt"
    #End If
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_FootnoteHarts.bas

```vb
Attribute VB_Name = "Rules_FootnoteHarts"
' ============================================================
' Rules_FootnoteHarts.bas
' Combined proofreading rules for footnotes per Hart's Rules:
'   - Rule24: flags documents that use endnotes instead of footnotes
'   - Rule25: every footnote should end with a full stop
'   - Rule26: footnotes should begin with a capital letter
'   - Rule27: flags unapproved footnote abbreviation variants
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

' ------------------------------------------------------------
'  Rule-name constants
' ------------------------------------------------------------
Private Const RULE24_NAME As String = "footnotes_not_endnotes"
Private Const RULE25_NAME As String = "footnote_terminal_full_stop"
Private Const RULE26_NAME As String = "footnote_initial_capital"
Private Const RULE27_NAME As String = "footnote_abbreviation_dictionary"

' ============================================================
'  RULE 24 -- FOOTNOTES NOT ENDNOTES
' ============================================================

Public Function Check_FootnotesNotEndnotes(doc As Document) As Collection
    Dim issues As New Collection
    Dim finding As Object

    On Error Resume Next

    Dim endCount As Long
    Dim fnCount As Long
    endCount = doc.Endnotes.Count
    fnCount = doc.Footnotes.Count

    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set Check_FootnotesNotEndnotes = issues
        Exit Function
    End If
    On Error GoTo 0

    If endCount > 0 And fnCount = 0 Then
        ' Document uses only endnotes
        Set finding = CreateIssueDict(RULE24_NAME, "document level", "Document uses endnotes instead of footnotes.", "Use footnotes rather than endnotes.", 0, 0, "error", False)
        issues.Add finding

    ElseIf endCount > 0 And fnCount > 0 Then
        ' Document uses both
        Set finding = CreateIssueDict(RULE24_NAME, "document level", "Document uses both footnotes and endnotes.", "Use footnotes rather than endnotes.", 0, 0, "error", False)
        issues.Add finding
    End If

    ' If only footnotes exist (endCount = 0): no finding

    Set Check_FootnotesNotEndnotes = issues
End Function

' ============================================================
'  RULE 25 -- FOOTNOTE TERMINAL FULL STOP
' ============================================================

Public Function Check_FootnoteTerminalFullStop(doc As Document) As Collection
    Dim issues As New Collection
    Dim fn As Footnote
    Dim finding As Object
    Dim locStr As String
    Dim noteText As String
    Dim trimmed As String
    Dim lastChar As String
    Dim penultChar As String
    Dim i As Long

    For i = 1 To doc.Footnotes.Count
        On Error Resume Next
        Set fn = doc.Footnotes(i)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote25
        End If
        On Error GoTo 0

        ' -- Check page range on the reference mark -----------
        On Error Resume Next
        If Not EngineIsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFootnote25
        End If
        On Error GoTo 0

        ' -- Get footnote text --------------------------------
        On Error Resume Next
        noteText = fn.Range.Text
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote25
        End If
        On Error GoTo 0

        ' -- Trim trailing whitespace / paragraph marks -------
        trimmed = noteText
        trimmed = TrimTrailingWhitespace(trimmed)

        ' -- Skip empty footnotes -----------------------------
        If Len(trimmed) = 0 Then GoTo NextFootnote25

        ' -- Get last character -------------------------------
        lastChar = Mid(trimmed, Len(trimmed), 1)

        ' -- If last char is closing bracket/quote, check penultimate --
        If IsClosingPunctuation(lastChar) Then
            If Len(trimmed) >= 2 Then
                penultChar = Mid(trimmed, Len(trimmed) - 1, 1)
                If penultChar = "." Then GoTo NextFootnote25
            End If
            ' Fall through to flag
        ElseIf lastChar = "." Then
            GoTo NextFootnote25
        End If

        ' -- Flag missing full stop ---------------------------
        On Error Resume Next
        locStr = EngineGetLocationString(fn.Reference, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
        On Error GoTo 0

        Set finding = CreateIssueDict(RULE25_NAME, locStr, "Footnote does not end with a full stop.", "Add a full stop at the end of the footnote.", fn.Range.Start, fn.Range.End, "warning", False)
        issues.Add finding

NextFootnote25:
    Next i

    Set Check_FootnoteTerminalFullStop = issues
End Function

' ============================================================
'  RULE 26 -- FOOTNOTE INITIAL CAPITAL
' ============================================================

Public Function Check_FootnoteInitialCapital(doc As Document) As Collection
    Dim issues As New Collection
    Dim allowed As Object
    Dim fn As Footnote
    Dim finding As Object
    Dim locStr As String
    Dim noteText As String
    Dim trimmed As String
    Dim token As String
    Dim firstCharCode As Long
    Dim i As Long
    Dim j As Long
    Dim ch As String

    ' -- Build allowed lower-case starts dictionary -----------
    Set allowed = CreateObject("Scripting.Dictionary")
    allowed.CompareMode = vbTextCompare
    allowed.Add "c", True
    allowed.Add "cf", True
    allowed.Add "cp", True
    allowed.Add "eg", True
    allowed.Add "ie", True
    allowed.Add "p", True
    allowed.Add "pp", True
    allowed.Add "ibid", True

    For i = 1 To doc.Footnotes.Count
        On Error Resume Next
        Set fn = doc.Footnotes(i)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote26
        End If
        On Error GoTo 0

        ' -- Check page range on the reference mark -----------
        On Error Resume Next
        If Not EngineIsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFootnote26
        End If
        On Error GoTo 0

        ' -- Get footnote text --------------------------------
        On Error Resume Next
        noteText = fn.Range.Text
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote26
        End If
        On Error GoTo 0

        ' -- Trim leading whitespace --------------------------
        trimmed = LTrim(noteText)
        If Len(trimmed) = 0 Then GoTo NextFootnote26

        ' -- Skip past leading punctuation (quotes, brackets) -
        j = 1
        Do While j <= Len(trimmed)
            ch = Mid(trimmed, j, 1)
            If IsLeadingPunctuation(ch) Then
                j = j + 1
            Else
                Exit Do
            End If
        Loop

        If j > Len(trimmed) Then GoTo NextFootnote26
        trimmed = Mid(trimmed, j)
        If Len(trimmed) = 0 Then GoTo NextFootnote26

        ' -- Extract first lexical token (letters only) -------
        token = ExtractFirstToken(trimmed)
        If Len(token) = 0 Then GoTo NextFootnote26

        ' -- Check if token is in allowed list ----------------
        If allowed.Exists(LCase(token)) Then GoTo NextFootnote26

        ' -- Check if first character is lower-case -----------
        firstCharCode = AscW(Mid(token, 1, 1))
        If firstCharCode >= 97 And firstCharCode <= 122 Then
            ' Lower-case and not in allowed list: flag
            On Error Resume Next
            locStr = EngineGetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE26_NAME, locStr, "Footnote begins with lower-case text outside the approved exceptions.", "Begin the footnote with a capital letter, unless it starts with an approved lower-case abbreviation.", fn.Range.Start, fn.Range.End, "warning", False)
            issues.Add finding
        End If

NextFootnote26:
    Next i

    Set Check_FootnoteInitialCapital = issues
End Function

' ============================================================
'  RULE 27 -- FOOTNOTE ABBREVIATION DICTIONARY
' ============================================================

Public Function Check_FootnoteAbbreviationDictionary(doc As Document) As Collection
    Dim issues As New Collection
    Dim approved As Object
    Dim approvedLC As Object
    Dim unapproved As Object
    Dim fn As Footnote
    Dim i As Long

    ' -- Build approved abbreviations set (case-sensitive) ----
    Set approved = CreateObject("Scripting.Dictionary")
    approved.CompareMode = vbBinaryCompare
    BuildApprovedDict approved

    ' -- Build approved lower-case set for dotted-form check --
    Set approvedLC = CreateObject("Scripting.Dictionary")
    approvedLC.CompareMode = vbTextCompare
    BuildApprovedLCDict approvedLC

    ' -- Build unapproved variant mapping (LCase key) --------
    Set unapproved = CreateObject("Scripting.Dictionary")
    unapproved.CompareMode = vbTextCompare
    BuildUnapprovedDict unapproved

    ' -- Process each footnote --------------------------------
    For i = 1 To doc.Footnotes.Count
        On Error Resume Next
        Set fn = doc.Footnotes(i)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFootnote27
        End If
        On Error GoTo 0

        ' -- Check page range on the reference mark -----------
        On Error Resume Next
        If Not EngineIsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFootnote27
        End If
        On Error GoTo 0

        CheckFootnoteText doc, fn, approved, approvedLC, unapproved, issues

NextFootnote27:
    Next i

    Set Check_FootnoteAbbreviationDictionary = issues
End Function

' ============================================================
'  PRIVATE HELPERS -- Rule 25
' ============================================================

' Strip trailing CR, LF, VT, and spaces
Private Function TrimTrailingWhitespace(ByVal s As String) As String
    Dim ch As String
    Do While Len(s) > 0
        ch = Mid(s, Len(s), 1)
        Select Case ch
            Case vbCr, vbLf, Chr(13), Chr(10), Chr(11), " ", vbTab
                s = Left(s, Len(s) - 1)
            Case Else
                Exit Do
        End Select
    Loop
    TrimTrailingWhitespace = s
End Function

' Check if character is a closing bracket or quote
Private Function IsClosingPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case ")", "]", ChrW(8217), ChrW(8221)
            IsClosingPunctuation = True
        Case Else
            IsClosingPunctuation = False
    End Select
End Function

' ============================================================
'  PRIVATE HELPERS -- Rule 26
' ============================================================

' Check if character is leading punctuation to skip
Private Function IsLeadingPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case "(", "[", ChrW(8216), ChrW(8220), """", "'"
            IsLeadingPunctuation = True
        Case Else
            IsLeadingPunctuation = False
    End Select
End Function

' Extract the first token of letters from a string
Private Function ExtractFirstToken(ByVal s As String) As String
    Dim i As Long
    Dim charCode As Long
    Dim result As String
    result = ""

    For i = 1 To Len(s)
        charCode = AscW(Mid(s, i, 1))
        ' A-Z = 65-90, a-z = 97-122
        If (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 97 And charCode <= 122) Then
            result = result & Mid(s, i, 1)
        Else
            Exit For
        End If
    Next i

    ExtractFirstToken = result
End Function

' ============================================================
'  PRIVATE HELPERS -- Rule 27
' ============================================================

' Check a single footnote's text for abbreviation issues
Private Sub CheckFootnoteText(doc As Document, _
                               fn As Footnote, _
                               ByRef approved As Object, _
                               ByRef approvedLC As Object, _
                               ByRef unapproved As Object, _
                               ByRef issues As Collection)
    Dim noteText As String
    Dim tokens() As String
    Dim token As String
    Dim stripped As String
    Dim noDots As String
    Dim lcToken As String
    Dim preferred As String
    Dim finding As Object
    Dim locStr As String
    Dim j As Long
    Dim issueText As String
    Dim suggText As String

    On Error Resume Next
    noteText = fn.Range.Text
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' -- Tokenize on spaces -----------------------------------
    tokens = Split(noteText, " ")

    For j = LBound(tokens) To UBound(tokens)
        token = Trim(tokens(j))
        If Len(token) = 0 Then GoTo NextToken

        ' Clean token boundaries: strip leading/trailing non-letter, non-dot chars
        token = CleanTokenBoundaries(token)
        If Len(token) = 0 Then GoTo NextToken

        ' -- Check 1: Unapproved variant (without trailing dot) --
        stripped = StripTrailingDot(token)
        lcToken = LCase(stripped)

        If unapproved.Exists(lcToken) Then
            preferred = unapproved(lcToken)

            On Error Resume Next
            locStr = EngineGetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            issueText = "Unapproved footnote abbreviation."
            suggText = "Use '" & preferred & "' instead of '" & stripped & "'."

            Set finding = CreateIssueDict(RULE27_NAME, locStr, issueText, suggText, fn.Range.Start, fn.Range.End, "warning", False)
            issues.Add finding
            GoTo NextToken
        End If

        ' -- Check 2: Dotted form of approved abbreviation -------
        ' Only flag tokens that contain dots
        If InStr(1, token, ".") > 0 Then
            ' Strip trailing dot and check
            stripped = StripTrailingDot(token)

            ' Remove all internal dots (e.g. "e.g." -> "eg", "i.e." -> "ie")
            noDots = Replace(stripped, ".", "")

            If Len(noDots) > 0 Then
                ' Check if the undotted form is an approved abbreviation
                If approvedLC.Exists(noDots) Then
                    ' This is a dotted form of an approved abbrev -- flag it
                    On Error Resume Next
                    locStr = EngineGetLocationString(fn.Reference, doc)
                    If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
                    On Error GoTo 0

                    issueText = "Unapproved footnote abbreviation."
                    suggText = "Use '" & noDots & "' instead of '" & token & "'."

                    Set finding = CreateIssueDict(RULE27_NAME, locStr, issueText, suggText, fn.Range.Start, fn.Range.End, "warning", False)
                    issues.Add finding
                    GoTo NextToken
                End If
            End If
        End If

NextToken:
    Next j
End Sub

' Build approved abbreviations dictionary (case-sensitive binary compare)
Private Sub BuildApprovedDict(ByRef d As Object)
    d.Add "Art", True
    d.Add "art", True
    d.Add "Arts", True
    d.Add "arts", True
    d.Add "ch", True
    d.Add "chs", True
    d.Add "c", True
    d.Add "cc", True
    d.Add "cl", True
    d.Add "cls", True
    d.Add "cp", True
    d.Add "cf", True
    d.Add "ed", True
    d.Add "eds", True
    d.Add "edn", True
    d.Add "edns", True
    d.Add "eg", True
    d.Add "etc", True
    d.Add "f", True
    d.Add "ff", True
    d.Add "fn", True
    d.Add "fns", True
    d.Add "ibid", True
    d.Add "ie", True
    d.Add "MS", True
    d.Add "MSS", True
    d.Add "n", True
    d.Add "nn", True
    d.Add "no", True
    d.Add "No", True
    d.Add "p", True
    d.Add "pp", True
    d.Add "para", True
    d.Add "paras", True
    d.Add "pt", True
    d.Add "reg", True
    d.Add "regs", True
    d.Add "r", True
    d.Add "rr", True
    d.Add "sch", True
    d.Add "s", True
    d.Add "ss", True
    d.Add "sub-s", True
    d.Add "sub-ss", True
    d.Add "trans", True
    d.Add "vol", True
    d.Add "vols", True
End Sub

' Build approved lower-case dictionary for dotted form checks (case-insensitive)
Private Sub BuildApprovedLCDict(ByRef d As Object)
    Dim abbrevs As Variant
    Dim k As Long
    abbrevs = Array("art", "arts", "ch", "chs", "c", "cc", "cl", "cls", _
                    "cp", "cf", "ed", "eds", "edn", "edns", "eg", "etc", _
                    "f", "ff", "fn", "fns", "ibid", "ie", "ms", "mss", _
                    "n", "nn", "no", "p", "pp", "para", "paras", "pt", _
                    "reg", "regs", "r", "rr", "sch", "s", "ss", _
                    "trans", "vol", "vols")
    For k = LBound(abbrevs) To UBound(abbrevs)
        If Not d.Exists(CStr(abbrevs(k))) Then
            d.Add CStr(abbrevs(k)), True
        End If
    Next k
End Sub

' Build unapproved variant mapping
Private Sub BuildUnapprovedDict(ByRef d As Object)
    d.Add "pgs", "pp"
    d.Add "sec", "s"
    d.Add "secs", "ss"
    d.Add "sect", "s"
    d.Add "sects", "ss"
    d.Add "para.", "para"
    d.Add "paras.", "paras"
End Sub

' Strip a single trailing dot from a token
Private Function StripTrailingDot(ByVal s As String) As String
    If Len(s) > 0 Then
        If Right(s, 1) = "." Then
            StripTrailingDot = Left(s, Len(s) - 1)
        Else
            StripTrailingDot = s
        End If
    Else
        StripTrailingDot = s
    End If
End Function

' Clean token boundaries -- strip leading/trailing characters
' that are not letters, digits, dots, or hyphens
Private Function CleanTokenBoundaries(ByVal s As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim ch As String
    Dim code As Long

    ' Strip leading non-word chars (keep letters, digits, dots, hyphens)
    startPos = 1
    Do While startPos <= Len(s)
        ch = Mid(s, startPos, 1)
        code = AscW(ch)
        If IsWordChar(code) Or ch = "." Or ch = "-" Then
            Exit Do
        End If
        startPos = startPos + 1
    Loop

    ' Strip trailing non-word chars (keep letters, digits, dots, hyphens)
    endPos = Len(s)
    Do While endPos >= startPos
        ch = Mid(s, endPos, 1)
        code = AscW(ch)
        If IsWordChar(code) Or ch = "." Or ch = "-" Then
            Exit Do
        End If
        endPos = endPos - 1
    Loop

    If startPos > endPos Then
        CleanTokenBoundaries = ""
    Else
        CleanTokenBoundaries = Mid(s, startPos, endPos - startPos + 1)
    End If
End Function

' Check if a character code is a letter or digit
Private Function IsWordChar(ByVal code As Long) As Boolean
    ' A-Z = 65-90, a-z = 97-122, 0-9 = 48-57
    IsWordChar = (code >= 65 And code <= 90) Or _
                 (code >= 97 And code <= 122) Or _
                 (code >= 48 And code <= 57)
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_FootnoteIntegrity.bas

```vb
Attribute VB_Name = "Rules_FootnoteIntegrity"
' ============================================================
' Rules_FootnoteIntegrity.bas
' Proofreading rule: checks footnote and endnote integrity.
'
' Checks performed:
'   1. Sequential numbering -- no gaps in index sequence
'   2. Placement after punctuation -- reference marks should
'      follow punctuation, not letters or spaces
'   3. Empty footnotes -- footnotes with no content
'   4. Duplicate content -- two footnotes with identical text
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "footnote_integrity"

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Function Check_FootnoteIntegrity(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Check footnotes -------------------------------------
    If doc.Footnotes.Count > 0 Then
        CheckNoteSequence doc, doc.Footnotes, "Footnote", issues
        CheckNotePlacement doc, doc.Footnotes, "Footnote", issues
        CheckEmptyNotes doc, doc.Footnotes, "Footnote", issues
        CheckDuplicateNotes doc, doc.Footnotes, "Footnote", issues
    End If

    ' -- Check endnotes --------------------------------------
    If doc.Endnotes.Count > 0 Then
        CheckEndnoteSequence doc, doc.Endnotes, "Endnote", issues
        CheckEndnotePlacement doc, doc.Endnotes, "Endnote", issues
        CheckEmptyEndnotes doc, doc.Endnotes, "Endnote", issues
        CheckDuplicateEndnotes doc, doc.Endnotes, "Endnote", issues
    End If

    Set Check_FootnoteIntegrity = issues
End Function

' ============================================================
'  PRIVATE: Check sequential numbering for footnotes
' ============================================================
Private Sub CheckNoteSequence(doc As Document, _
                               notes As Footnotes, _
                               noteType As String, _
                               ByRef issues As Collection)
    Dim i As Long
    Dim expectedIdx As Long
    Dim fn As Footnote
    Dim finding As Object
    Dim locStr As String

    expectedIdx = 1

    For i = 1 To notes.Count
        Set fn = notes(i)

        On Error Resume Next
        If Not EngineIsInPageRange(fn.Reference) Then
            expectedIdx = expectedIdx + 1
            On Error GoTo 0
            GoTo NextFootnote
        End If
        On Error GoTo 0

        If fn.Index <> expectedIdx Then
            On Error Resume Next
            locStr = EngineGetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME, locStr, noteType & " numbering gap: expected " & expectedIdx & ", found " & fn.Index, "Renumber " & LCase(noteType) & "s sequentially", fn.Reference.Start, fn.Reference.End, "error")
            issues.Add finding
        End If

        expectedIdx = expectedIdx + 1

NextFootnote:
    Next i
End Sub

' ============================================================
'  PRIVATE: Check sequential numbering for endnotes
' ============================================================
Private Sub CheckEndnoteSequence(doc As Document, _
                                  notes As Endnotes, _
                                  noteType As String, _
                                  ByRef issues As Collection)
    Dim i As Long
    Dim expectedIdx As Long
    Dim en As Endnote
    Dim finding As Object
    Dim locStr As String

    expectedIdx = 1

    For i = 1 To notes.Count
        Set en = notes(i)

        On Error Resume Next
        If Not EngineIsInPageRange(en.Reference) Then
            expectedIdx = expectedIdx + 1
            On Error GoTo 0
            GoTo NextEndnoteSeq
        End If
        On Error GoTo 0

        If en.Index <> expectedIdx Then
            On Error Resume Next
            locStr = EngineGetLocationString(en.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME, locStr, noteType & " numbering gap: expected " & expectedIdx & ", found " & en.Index, "Renumber " & LCase(noteType) & "s sequentially", en.Reference.Start, en.Reference.End, "error")
            issues.Add finding
        End If

        expectedIdx = expectedIdx + 1

NextEndnoteSeq:
    Next i
End Sub

' ============================================================
'  PRIVATE: Check placement after punctuation for footnotes
' ============================================================
Private Sub CheckNotePlacement(doc As Document, _
                                notes As Footnotes, _
                                noteType As String, _
                                ByRef issues As Collection)
    Dim i As Long
    Dim fn As Footnote
    Dim finding As Object
    Dim locStr As String
    Dim charBefore As String
    Dim refStart As Long

    For i = 1 To notes.Count
        Set fn = notes(i)

        On Error Resume Next
        If Not EngineIsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFnPlace
        End If
        On Error GoTo 0

        refStart = fn.Reference.Start

        ' Check character before the reference mark
        If refStart > 0 Then
            On Error Resume Next
            charBefore = doc.Range(refStart - 1, refStart).Text
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo NextFnPlace
            End If
            On Error GoTo 0

            If Not IsPunctuation(charBefore) Then
                On Error Resume Next
                locStr = EngineGetLocationString(fn.Reference, doc)
                If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
                On Error GoTo 0

                Set finding = CreateIssueDict(RULE_NAME, locStr, noteType & " " & fn.Index & " reference not placed after punctuation", "Place " & LCase(noteType) & " reference after punctuation mark", fn.Reference.Start, fn.Reference.End, "error")
                issues.Add finding
            End If
        End If

NextFnPlace:
    Next i
End Sub

' ============================================================
'  PRIVATE: Check placement after punctuation for endnotes
' ============================================================
Private Sub CheckEndnotePlacement(doc As Document, _
                                   notes As Endnotes, _
                                   noteType As String, _
                                   ByRef issues As Collection)
    Dim i As Long
    Dim en As Endnote
    Dim finding As Object
    Dim locStr As String
    Dim charBefore As String
    Dim refStart As Long

    For i = 1 To notes.Count
        Set en = notes(i)

        On Error Resume Next
        If Not EngineIsInPageRange(en.Reference) Then
            On Error GoTo 0
            GoTo NextEnPlace
        End If
        On Error GoTo 0

        refStart = en.Reference.Start

        If refStart > 0 Then
            On Error Resume Next
            charBefore = doc.Range(refStart - 1, refStart).Text
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo NextEnPlace
            End If
            On Error GoTo 0

            If Not IsPunctuation(charBefore) Then
                On Error Resume Next
                locStr = EngineGetLocationString(en.Reference, doc)
                If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
                On Error GoTo 0

                Set finding = CreateIssueDict(RULE_NAME, locStr, noteType & " " & en.Index & " reference not placed after punctuation", "Place " & LCase(noteType) & " reference after punctuation mark", en.Reference.Start, en.Reference.End, "error")
                issues.Add finding
            End If
        End If

NextEnPlace:
    Next i
End Sub

' ============================================================
'  PRIVATE: Check for empty footnotes
' ============================================================
Private Sub CheckEmptyNotes(doc As Document, _
                             notes As Footnotes, _
                             noteType As String, _
                             ByRef issues As Collection)
    Dim i As Long
    Dim fn As Footnote
    Dim finding As Object
    Dim locStr As String
    Dim noteText As String

    For i = 1 To notes.Count
        Set fn = notes(i)

        On Error Resume Next
        If Not EngineIsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFnEmpty
        End If
        On Error GoTo 0

        On Error Resume Next
        noteText = fn.Range.Text
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextFnEmpty
        On Error GoTo 0

        noteText = Trim(Replace(noteText, vbCr, ""))
        noteText = Trim(Replace(noteText, vbLf, ""))

        If Len(noteText) = 0 Then
            On Error Resume Next
            locStr = EngineGetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME, locStr, noteType & " " & fn.Index & " has empty content", "Add content or remove the empty " & LCase(noteType), fn.Reference.Start, fn.Reference.End, "error")
            issues.Add finding
        End If

NextFnEmpty:
    Next i
End Sub

' ============================================================
'  PRIVATE: Check for empty endnotes
' ============================================================
Private Sub CheckEmptyEndnotes(doc As Document, _
                                notes As Endnotes, _
                                noteType As String, _
                                ByRef issues As Collection)
    Dim i As Long
    Dim en As Endnote
    Dim finding As Object
    Dim locStr As String
    Dim noteText As String

    For i = 1 To notes.Count
        Set en = notes(i)

        On Error Resume Next
        If Not EngineIsInPageRange(en.Reference) Then
            On Error GoTo 0
            GoTo NextEnEmpty
        End If
        On Error GoTo 0

        On Error Resume Next
        noteText = en.Range.Text
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextEnEmpty
        On Error GoTo 0

        noteText = Trim(Replace(noteText, vbCr, ""))
        noteText = Trim(Replace(noteText, vbLf, ""))

        If Len(noteText) = 0 Then
            On Error Resume Next
            locStr = EngineGetLocationString(en.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME, locStr, noteType & " " & en.Index & " has empty content", "Add content or remove the empty " & LCase(noteType), en.Reference.Start, en.Reference.End, "error")
            issues.Add finding
        End If

NextEnEmpty:
    Next i
End Sub

' ============================================================
'  PRIVATE: Check for duplicate footnote content
' ============================================================
Private Sub CheckDuplicateNotes(doc As Document, _
                                 notes As Footnotes, _
                                 noteType As String, _
                                 ByRef issues As Collection)
    Dim contentDict As Object
    Set contentDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim fn As Footnote
    Dim finding As Object
    Dim locStr As String
    Dim noteText As String
    Dim cleanText As String

    For i = 1 To notes.Count
        Set fn = notes(i)

        On Error Resume Next
        If Not EngineIsInPageRange(fn.Reference) Then
            On Error GoTo 0
            GoTo NextFnDup
        End If
        On Error GoTo 0

        On Error Resume Next
        noteText = fn.Range.Text
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextFnDup
        On Error GoTo 0

        cleanText = Trim(Replace(noteText, vbCr, ""))
        cleanText = Trim(Replace(cleanText, vbLf, ""))

        ' Skip empty notes (already flagged separately)
        If Len(cleanText) = 0 Then GoTo NextFnDup

        If contentDict.Exists(cleanText) Then
            ' This is a duplicate
            Dim firstIdx As Long
            firstIdx = CLng(contentDict(cleanText))

            On Error Resume Next
            locStr = EngineGetLocationString(fn.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME, locStr, noteType & " " & fn.Index & " has identical content to " & LCase(noteType) & " " & firstIdx, "Remove duplicate or differentiate content", fn.Reference.Start, fn.Reference.End, "possible_error")
            issues.Add finding
        Else
            contentDict.Add cleanText, fn.Index
        End If

NextFnDup:
    Next i
End Sub

' ============================================================
'  PRIVATE: Check for duplicate endnote content
' ============================================================
Private Sub CheckDuplicateEndnotes(doc As Document, _
                                    notes As Endnotes, _
                                    noteType As String, _
                                    ByRef issues As Collection)
    Dim contentDict As Object
    Set contentDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim en As Endnote
    Dim finding As Object
    Dim locStr As String
    Dim noteText As String
    Dim cleanText As String

    For i = 1 To notes.Count
        Set en = notes(i)

        On Error Resume Next
        If Not EngineIsInPageRange(en.Reference) Then
            On Error GoTo 0
            GoTo NextEnDup
        End If
        On Error GoTo 0

        On Error Resume Next
        noteText = en.Range.Text
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextEnDup
        On Error GoTo 0

        cleanText = Trim(Replace(noteText, vbCr, ""))
        cleanText = Trim(Replace(cleanText, vbLf, ""))

        If Len(cleanText) = 0 Then GoTo NextEnDup

        If contentDict.Exists(cleanText) Then
            Dim firstEnIdx As Long
            firstEnIdx = CLng(contentDict(cleanText))

            On Error Resume Next
            locStr = EngineGetLocationString(en.Reference, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME, locStr, noteType & " " & en.Index & " has identical content to " & LCase(noteType) & " " & firstEnIdx, "Remove duplicate or differentiate content", en.Reference.Start, en.Reference.End, "possible_error")
            issues.Add finding
        Else
            contentDict.Add cleanText, en.Index
        End If

NextEnDup:
    Next i
End Sub

' ============================================================
'  PRIVATE: Check if character is punctuation
' ============================================================
Private Function IsPunctuation(ByVal ch As String) As Boolean
    Select Case ch
        Case ".", ",", ";", ":", """", "'", ")", _
             ChrW(8221), ChrW(8217), ChrW(8220), ChrW(8216), _
             "!", "?"
            IsPunctuation = True
        Case Else
            IsPunctuation = False
    End Select
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_Formatting.bas

```vb
Attribute VB_Name = "Rules_Formatting"
' ============================================================
' Rules_Formatting.bas
' Combined module for formatting-related rules:
'   - Rule06: Paragraph break consistency (headings)
'   - Rule11: Font consistency (headings, body, footnotes)
'
' IsBlockQuotePara is a public helper used by other modules.
' It requires STRONG indicators beyond mere indentation:
'   - Quote-related style name (definitive)
'   - Indentation + quotation-mark wrapping
'   - Indentation + entirely italic text
' Indentation + smaller font alone is NOT sufficient.
' ============================================================
Option Explicit

Private Const RULE_NAME_PARAGRAPH_BREAK As String = "paragraph_break_consistency"
Private Const RULE_NAME_FONT As String = "font_consistency"

' ============================================================
'  RULE 06 HELPERS
' ============================================================

' -- Classify spacing pattern after a heading ----------------
' Returns: "no_spacing", "spacing_Npt", or "manual_double_break"
Private Function ClassifyAfterSpacing(para As Paragraph, doc As Document, paraIdx As Long) As String
    Dim spAfter As Single
    spAfter = para.Format.SpaceAfter

    ' Check if the next paragraph is empty (manual double break)
    Dim totalParas As Long
    totalParas = doc.Paragraphs.Count
    If paraIdx < totalParas Then
        Dim nextPara As Paragraph
        Set nextPara = doc.Paragraphs(paraIdx + 1)
        Dim nextText As String
        nextText = nextPara.Range.Text
        ' An empty paragraph contains only vbCr
        If nextText = vbCr Then
            ClassifyAfterSpacing = "manual_double_break"
            Exit Function
        End If
    End If

    If spAfter = 0 Then
        ClassifyAfterSpacing = "no_spacing"
    Else
        ClassifyAfterSpacing = "spacing_" & CStr(CLng(spAfter)) & "pt"
    End If
End Function

' -- Classify SpaceBefore pattern ----------------------------
Private Function ClassifyBeforeSpacing(para As Paragraph) As String
    Dim spBefore As Single
    spBefore = para.Format.SpaceBefore
    If spBefore = 0 Then
        ClassifyBeforeSpacing = "before_0pt"
    Else
        ClassifyBeforeSpacing = "before_" & CStr(CLng(spBefore)) & "pt"
    End If
End Function

' ============================================================
'  RULE 11 HELPERS
' ============================================================

' -- Helper: build a font profile key ------------------------
Private Function FontKey(ByVal fontName As String, ByVal fontSize As Single) As String
    FontKey = fontName & "|" & CStr(fontSize)
End Function

' -- Helper: find dominant key in a dictionary of counts -----
Private Function GetDominant(counts As Object) As String
    Dim k As Variant
    Dim maxCnt As Long
    Dim domKey As String
    maxCnt = 0
    domKey = ""
    For Each k In counts.keys
        If counts(k) > maxCnt Then
            maxCnt = counts(k)
            domKey = CStr(k)
        End If
    Next k
    GetDominant = domKey
End Function

' -- Helper: parse font key back to readable description -----
' ------------------------------------------------------------
'  PUBLIC: Detect block quote / indented extract paragraphs.
'
'  STRICT RULE: Indentation alone is NEVER enough.
'  Smaller font + indentation alone is NEVER enough.
'  A block quote must have at least one of:
'    1. A block-quote style (name contains "quote"/"block"/"extract")
'    2. Enclosing quotation marks AND indentation
'    3. Entirely italic text AND indentation
'  Lists, numbered paragraphs, and bullet items are explicitly excluded.
' ------------------------------------------------------------
Public Function IsBlockQuotePara(para As Paragraph) As Boolean
    IsBlockQuotePara = False
    On Error Resume Next

    ' ==========================================================
    '  CHECK 0: Exclude list paragraphs (numbered, bulleted, etc.)
    '  Lists must NEVER be treated as block quotes.
    ' ==========================================================
    Dim listLvl As Long
    listLvl = 0
    listLvl = para.Range.ListFormat.ListLevelNumber
    If Err.Number <> 0 Then listLvl = 0: Err.Clear
    ' ListLevelNumber > 0 means this paragraph is in a list
    If listLvl > 0 Then
        On Error GoTo 0
        Exit Function
    End If

    ' Also check for list-like text patterns (manual numbering)
    Dim pTextRaw As String
    pTextRaw = ""
    pTextRaw = para.Range.Text
    If Err.Number <> 0 Then pTextRaw = "": Err.Clear
    On Error GoTo 0
    Dim pTextTrimmed As String
    pTextTrimmed = Replace(Replace(Replace(pTextRaw, vbCr, ""), vbTab, ""), ChrW(160), " ")
    pTextTrimmed = Trim$(pTextTrimmed)

    ' Check for bullet-like or number-list-like starts
    If Len(pTextTrimmed) > 1 Then
        Dim firstTwo As String
        firstTwo = Left$(pTextTrimmed, 2)
        ' Bullet characters: bullet, en-dash, em-dash, hyphen
        If Left$(pTextTrimmed, 1) = ChrW(8226) Or _
           Left$(pTextTrimmed, 1) = ChrW(8211) & " " Or _
           firstTwo = "- " Or firstTwo = "* " Then
            On Error GoTo 0
            Exit Function
        End If
        ' Numbered list pattern: "(a)", "(i)", "(1)", "1.", "a.", "i."
        If pTextTrimmed Like "(#)*" Or pTextTrimmed Like "(##)*" Or _
           pTextTrimmed Like "([a-z])*" Or pTextTrimmed Like "([ivx])*" Or _
           pTextTrimmed Like "#.*" Or pTextTrimmed Like "##.*" Or _
           pTextTrimmed Like "[a-z].*" Then
            On Error GoTo 0
            Exit Function
        End If
    End If

    ' Also check ListFormat.ListString for auto-numbered lists
    Dim listStr As String
    listStr = ""
    On Error Resume Next
    listStr = para.Range.ListFormat.ListString
    If Err.Number <> 0 Then listStr = "": Err.Clear
    If Len(listStr) > 0 Then
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  CHECK 1: Style name for quote/block/extract keywords
    '  (Definitive indicator - no other checks needed)
    ' ==========================================================
    Dim sn As String
    sn = LCase(para.Style.NameLocal)
    If Err.Number <> 0 Then sn = "": Err.Clear
    If InStr(sn, "quote") > 0 Or InStr(sn, "block") > 0 Or _
       InStr(sn, "extract") > 0 Then
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  INDENTATION CHECK
    '  All remaining indicators require indentation.
    ' ==========================================================
    Dim leftInd As Single
    leftInd = para.Format.LeftIndent
    If Err.Number <> 0 Then leftInd = 0: Err.Clear
    On Error GoTo 0

    ' No indentation = not a block quote (style check already done above)
    If leftInd <= 18 Then
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  CHECK 2: Indentation + quotation marks wrapping
    '  Starts or ends with a quotation mark character.
    ' ==========================================================
    Dim startsWithQuote As Boolean
    Dim endsWithQuote As Boolean
    startsWithQuote = False
    endsWithQuote = False
    If Len(pTextTrimmed) > 1 Then
        Dim fcChar As String
        Dim lcChar As String
        fcChar = Left$(pTextTrimmed, 1)
        lcChar = Right$(pTextTrimmed, 1)
        startsWithQuote = (fcChar = Chr(34) Or fcChar = ChrW(8220) Or fcChar = ChrW(8216))
        endsWithQuote = (lcChar = Chr(34) Or lcChar = ChrW(8221) Or lcChar = ChrW(8217))
    End If

    ' Block quote if indented AND wrapped in quotation marks
    If startsWithQuote Or endsWithQuote Then
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  CHECK 3: Indentation + entirely italic
    '  wdTrue (-1) means ALL text in the range is italic.
    ' ==========================================================
    Dim italVal As Long
    On Error Resume Next
    italVal = para.Range.Font.Italic
    If Err.Number <> 0 Then italVal = 0: Err.Clear
    If italVal = -1 Then  ' wdTrue = -1 means ALL italic
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  DEFAULT: Indented but no strong indicator = NOT a block quote.
    '  Smaller font + indentation alone is deliberately NOT enough.
    '  This prevents indented lists, definitions, and body text
    '  from being misclassified.
    ' ==========================================================

    On Error GoTo 0
End Function

Private Function FontDescription(ByVal fKey As String) As String
    Dim parts() As String
    parts = Split(fKey, "|")
    If UBound(parts) >= 1 Then
        FontDescription = parts(0) & " " & parts(1) & "pt"
    Else
        FontDescription = fKey
    End If
End Function

' ============================================================
'  RULE 06: PARAGRAPH BREAK CONSISTENCY
' ============================================================
Public Function Check_ParagraphBreakConsistency(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim lvl As Long
    Dim info() As Variant

    On Error Resume Next

    ' -- Dictionaries keyed by outline level -----------------
    ' afterPatterns:  level -> Dictionary(pattern -> count)
    ' beforePatterns: level -> Dictionary(pattern -> count)
    ' headingInfos:   level -> Collection of Array(paraIdx, afterPattern, beforePattern, rangeStart, rangeEnd, text)
    Dim afterPatterns As Object
    Set afterPatterns = CreateObject("Scripting.Dictionary")
    Dim beforePatterns As Object
    Set beforePatterns = CreateObject("Scripting.Dictionary")
    Dim headingInfos As Object
    Set headingInfos = CreateObject("Scripting.Dictionary")

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        lvl = para.OutlineLevel
        If lvl < wdOutlineLevel1 Or lvl > wdOutlineLevel9 Then GoTo NextPara

        ' Page range filter
        If Not EngineIsInPageRange(para.Range) Then GoTo NextPara

        ' Classify after-spacing
        Dim aftPat As String
        aftPat = ClassifyAfterSpacing(para, doc, paraIdx)

        ' Classify before-spacing
        Dim befPat As String
        befPat = ClassifyBeforeSpacing(para)

        ' -- Track after-spacing counts ---------------------
        If Not afterPatterns.Exists(lvl) Then
            afterPatterns.Add lvl, CreateObject("Scripting.Dictionary")
        End If
        Dim aftDict As Object
        Set aftDict = afterPatterns(lvl)
        If aftDict.Exists(aftPat) Then
            aftDict(aftPat) = aftDict(aftPat) + 1
        Else
            aftDict.Add aftPat, 1
        End If

        ' -- Track before-spacing counts --------------------
        If Not beforePatterns.Exists(lvl) Then
            beforePatterns.Add lvl, CreateObject("Scripting.Dictionary")
        End If
        Dim befDict As Object
        Set befDict = beforePatterns(lvl)
        If befDict.Exists(befPat) Then
            befDict(befPat) = befDict(befPat) + 1
        Else
            befDict.Add befPat, 1
        End If

        ' -- Store heading info -----------------------------
        If Not headingInfos.Exists(lvl) Then
            headingInfos.Add lvl, New Collection
        End If
        ReDim info(0 To 5)
        info(0) = paraIdx
        info(1) = aftPat
        info(2) = befPat
        info(3) = para.Range.Start
        info(4) = para.Range.End
        info(5) = Trim$(Replace(para.Range.Text, vbCr, ""))
        headingInfos(lvl).Add info
NextPara:
    Next para

    ' -- Determine dominant patterns and flag deviations -----
    Dim lvlKey As Variant
    For Each lvlKey In headingInfos.keys
        Dim hdgs As Collection
        Set hdgs = headingInfos(lvlKey)
        If hdgs.Count <= 1 Then GoTo NextLevel

        ' Find dominant after-pattern
        Dim domAfter As String
        domAfter = ""
        Dim maxCnt As Long
        maxCnt = 0
        If afterPatterns.Exists(lvlKey) Then
            Set aftDict = afterPatterns(lvlKey)
            Dim pk As Variant
            For Each pk In aftDict.keys
                If aftDict(pk) > maxCnt Then
                    maxCnt = aftDict(pk)
                    domAfter = CStr(pk)
                End If
            Next pk
        End If

        ' Find dominant before-pattern
        Dim domBefore As String
        domBefore = ""
        maxCnt = 0
        If beforePatterns.Exists(lvlKey) Then
            Set befDict = beforePatterns(lvlKey)
            For Each pk In befDict.keys
                If befDict(pk) > maxCnt Then
                    maxCnt = befDict(pk)
                    domBefore = CStr(pk)
                End If
            Next pk
        End If

        ' Flag outliers
        Dim h As Long
        For h = 1 To hdgs.Count
            Dim hInfo As Variant
            hInfo = hdgs(h)

            Dim hAft As String
            hAft = CStr(hInfo(1))
            Dim hBef As String
            hBef = CStr(hInfo(2))
            Dim hText As String
            hText = CStr(hInfo(5))

            ' Check after-spacing deviation
            If hAft <> domAfter And Len(domAfter) > 0 Then
                Dim findingA As Object
                Dim rngA As Range
                Set rngA = doc.Range(CLng(hInfo(3)), CLng(hInfo(4)))
                Dim locA As String
                locA = EngineGetLocationString(rngA, doc)

                Set findingA = CreateIssueDict(RULE_NAME_PARAGRAPH_BREAK, locA, "After-heading spacing inconsistency at '" & hText & "': uses " & hAft & " but dominant pattern for level " & CLng(lvlKey) & " headings is " & domAfter, "Change spacing after this heading to match: " & domAfter, CLng(hInfo(3)), CLng(hInfo(4)), "possible_error")
                issues.Add findingA
            End If

            ' Check before-spacing deviation
            If hBef <> domBefore And Len(domBefore) > 0 Then
                Dim findingB As Object
                Dim rngB As Range
                Set rngB = doc.Range(CLng(hInfo(3)), CLng(hInfo(4)))
                Dim locB As String
                locB = EngineGetLocationString(rngB, doc)

                Set findingB = CreateIssueDict(RULE_NAME_PARAGRAPH_BREAK, locB, "Before-heading spacing inconsistency at '" & hText & "': uses " & hBef & " but dominant pattern for level " & CLng(lvlKey) & " headings is " & domBefore, "Change spacing before this heading to match: " & domBefore, CLng(hInfo(3)), CLng(hInfo(4)), "possible_error")
                issues.Add findingB
            End If
        Next h
NextLevel:
    Next lvlKey

    If Err.Number <> 0 Then
        Debug.Print "Check_ParagraphBreakConsistency: exiting with Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    Set Check_ParagraphBreakConsistency = issues
End Function

' ============================================================
'  RULE 11: FONT CONSISTENCY
'  Type-based approach: classify paragraphs into heading,
'  body, block-quote; compute dominant font per type;
'  flag outliers within each type.
'
'  Block-quote classification in font consistency uses the
'  same strict criteria as IsBlockQuotePara: indentation
'  plus italic, or indentation plus quotation wrapping.
'  Indentation + smaller font alone does NOT classify as
'  block quote here either.
' ============================================================
Public Function Check_FontConsistency(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' ==========================================================
    '  SINGLE MERGED PASS: Classify paragraphs and collect font
    '  tallies in one scan.
    ' ==========================================================
    Dim bodyIndents As Object   ' LeftIndent (rounded) -> count
    Set bodyIndents = CreateObject("Scripting.Dictionary")
    Dim bodySizes As Object     ' FontSize (rounded) -> count
    Set bodySizes = CreateObject("Scripting.Dictionary")

    Dim headingFonts As Object
    Set headingFonts = CreateObject("Scripting.Dictionary")
    Dim bodyFonts As Object
    Set bodyFonts = CreateObject("Scripting.Dictionary")
    Dim bqFonts As Object
    Set bqFonts = CreateObject("Scripting.Dictionary")
    Dim footnoteFonts As Object
    Set footnoteFonts = CreateObject("Scripting.Dictionary")

    ' Cache paragraph metadata in arrays to avoid re-scanning
    Dim paraCap As Long
    paraCap = 512
    Dim pLevels() As Long       ' outline level
    Dim pIndents() As Single    ' left indent
    Dim pFontNames() As String  ' font name
    Dim pFontSizes() As Single  ' font size
    Dim pStarts() As Long       ' range start
    Dim pEnds() As Long         ' range end
    Dim pTypes() As String      ' "heading"/"body"/"block_quote"/""
    Dim pInRange() As Boolean   ' in page range
    ReDim pLevels(0 To paraCap - 1)
    ReDim pIndents(0 To paraCap - 1)
    ReDim pFontNames(0 To paraCap - 1)
    ReDim pFontSizes(0 To paraCap - 1)
    ReDim pStarts(0 To paraCap - 1)
    ReDim pEnds(0 To paraCap - 1)
    ReDim pTypes(0 To paraCap - 1)
    ReDim pInRange(0 To paraCap - 1)

    Dim para As Paragraph
    Dim paraIdx As Long
    Dim fk As String

    ' -- Single scan: collect all paragraph metadata --
    paraIdx = 0
    For Each para In doc.Paragraphs
        ' Grow arrays if needed
        If paraIdx >= paraCap Then
            paraCap = paraCap * 2
            ReDim Preserve pLevels(0 To paraCap - 1)
            ReDim Preserve pIndents(0 To paraCap - 1)
            ReDim Preserve pFontNames(0 To paraCap - 1)
            ReDim Preserve pFontSizes(0 To paraCap - 1)
            ReDim Preserve pStarts(0 To paraCap - 1)
            ReDim Preserve pEnds(0 To paraCap - 1)
            ReDim Preserve pTypes(0 To paraCap - 1)
            ReDim Preserve pInRange(0 To paraCap - 1)
        End If

        pTypes(paraIdx) = ""
        pInRange(paraIdx) = EngineIsInPageRange(para.Range)
        If Not pInRange(paraIdx) Then
            paraIdx = paraIdx + 1
            GoTo NextScanPara
        End If

        Dim lvl As Long
        lvl = para.OutlineLevel
        If Err.Number <> 0 Then lvl = wdOutlineLevelBodyText: Err.Clear
        pLevels(paraIdx) = lvl

        Dim curInd As Single
        curInd = para.Format.LeftIndent
        If Err.Number <> 0 Then curInd = 0: Err.Clear
        pIndents(paraIdx) = curInd

        pStarts(paraIdx) = para.Range.Start
        pEnds(paraIdx) = para.Range.End

        ' Font info (read once, cache for reuse)
        Dim curFontName As String
        Dim curFontSize As Single
        curFontName = para.Range.Font.Name
        If Err.Number <> 0 Then curFontName = "": Err.Clear
        curFontSize = para.Range.Font.Size
        If Err.Number <> 0 Then curFontSize = 0: Err.Clear
        If curFontSize > 1000 Then curFontSize = 0
        pFontNames(paraIdx) = curFontName
        pFontSizes(paraIdx) = curFontSize

        ' Tally body-text indent and size (for dominant calculation)
        If lvl = wdOutlineLevelBodyText And curFontSize > 0 Then
            Dim indKey As String
            indKey = CStr(CLng(curInd))
            If bodyIndents.Exists(indKey) Then
                bodyIndents(indKey) = bodyIndents(indKey) + 1
            Else
                bodyIndents.Add indKey, 1
            End If
            Dim szKey As String
            szKey = CStr(CLng(curFontSize * 10))
            If bodySizes.Exists(szKey) Then
                bodySizes(szKey) = bodySizes(szKey) + 1
            Else
                bodySizes.Add szKey, 1
            End If
        End If

        paraIdx = paraIdx + 1
NextScanPara:
    Next para
    Dim totalParas As Long
    totalParas = paraIdx

    ' Determine dominant body indent and font size
    Dim domBodyIndent As Single
    Dim domBodySizeTenths As Long
    Dim domBodySize As Single
    Dim tmpDomKey As String
    tmpDomKey = GetDominant(bodyIndents)
    If Len(tmpDomKey) > 0 Then domBodyIndent = CSng(tmpDomKey) Else domBodyIndent = 0
    tmpDomKey = GetDominant(bodySizes)
    If Len(tmpDomKey) > 0 Then domBodySizeTenths = CLng(tmpDomKey) Else domBodySizeTenths = 0
    domBodySize = CSng(domBodySizeTenths) / 10#

    ' -- Classify paragraphs and tally fonts (in memory, no doc access) --
    ' Block-quote classification uses STRICT criteria:
    '   - Indentation + full italic (font.Italic = wdTrue)
    '   - Indentation + quote wrapping (first/last char is quote mark)
    '   - Style name with "quote"/"block"/"extract"
    ' Indentation + smaller font alone = classified as body, NOT block_quote.
    Dim pi As Long
    For pi = 0 To totalParas - 1
        If Not pInRange(pi) Then GoTo NextClassify

        Dim paraType As String
        paraType = ""
        Dim isHeading As Boolean
        isHeading = (pLevels(pi) >= wdOutlineLevel1 And pLevels(pi) <= wdOutlineLevel9)

        If isHeading Then
            paraType = "heading"
        ElseIf pLevels(pi) = wdOutlineLevelBodyText Then
            ' Use IsBlockQuotePara for strict classification
            ' This requires doc access but is per-paragraph, not per-run
            Dim bqPara As Paragraph
            Set bqPara = Nothing
            Dim bqRng As Range
            Set bqRng = doc.Range(pStarts(pi), pEnds(pi))
            If Err.Number = 0 Then
                ' Try to get the paragraph object
                ' Use a lightweight check: if IsBlockQuotePara was already
                ' determined during the scan, use it.  Otherwise classify
                ' conservatively as body.
                Dim isBQ As Boolean
                isBQ = False

                ' Check style name
                Dim paraStyleName As String
                paraStyleName = ""
                paraStyleName = bqRng.ParagraphStyle
                If Err.Number <> 0 Then paraStyleName = "": Err.Clear
                Dim lsn As String
                lsn = LCase$(paraStyleName)
                If InStr(lsn, "quote") > 0 Or InStr(lsn, "block") > 0 Or _
                   InStr(lsn, "extract") > 0 Then
                    isBQ = True
                End If

                ' Check indentation + italic
                If Not isBQ And pIndents(pi) > domBodyIndent + 18 Then
                    Dim italCheck As Long
                    italCheck = bqRng.Font.Italic
                    If Err.Number <> 0 Then italCheck = 0: Err.Clear
                    If italCheck = -1 Then isBQ = True  ' wdTrue = all italic
                End If

                ' Check indentation + quotation wrapping
                If Not isBQ And pIndents(pi) > domBodyIndent + 18 Then
                    Dim bqText As String
                    bqText = ""
                    bqText = bqRng.Text
                    If Err.Number <> 0 Then bqText = "": Err.Clear
                    bqText = Trim$(Replace(Replace(bqText, vbCr, ""), vbTab, ""))
                    If Len(bqText) > 1 Then
                        Dim bqFirst As String
                        Dim bqLast As String
                        bqFirst = Left$(bqText, 1)
                        bqLast = Right$(bqText, 1)
                        If bqFirst = Chr(34) Or bqFirst = ChrW(8220) Or bqFirst = ChrW(8216) Or _
                           bqLast = Chr(34) Or bqLast = ChrW(8221) Or bqLast = ChrW(8217) Then
                            isBQ = True
                        End If
                    End If
                End If

                If isBQ Then
                    paraType = "block_quote"
                Else
                    paraType = "body"
                End If
            Else
                Err.Clear
                paraType = "body"
            End If
        End If

        pTypes(pi) = paraType

        ' Tally font for this type
        If Len(pFontNames(pi)) > 0 And pFontSizes(pi) > 0 Then
            fk = FontKey(pFontNames(pi), pFontSizes(pi))
            Select Case paraType
                Case "heading"
                    If headingFonts.Exists(fk) Then
                        headingFonts(fk) = headingFonts(fk) + 1
                    Else
                        headingFonts.Add fk, 1
                    End If
                Case "body"
                    If bodyFonts.Exists(fk) Then
                        bodyFonts(fk) = bodyFonts(fk) + 1
                    Else
                        bodyFonts.Add fk, 1
                    End If
                Case "block_quote"
                    If bqFonts.Exists(fk) Then
                        bqFonts(fk) = bqFonts(fk) + 1
                    Else
                        bqFonts.Add fk, 1
                    End If
            End Select
        End If
NextClassify:
    Next pi

    ' -- Footnotes ------------------------------------------
    Dim fn As Footnote
    For Each fn In doc.Footnotes
        If Not EngineIsInPageRange(fn.Range) Then GoTo NextFootnote

        Dim fnFontName As String
        Dim fnFontSize As Single
        fnFontName = fn.Range.Font.Name
        fnFontSize = fn.Range.Font.Size

        If Len(fnFontName) > 0 And fnFontSize > 0 And fnFontSize < 1000 Then
            fk = FontKey(fnFontName, fnFontSize)
            If footnoteFonts.Exists(fk) Then
                footnoteFonts(fk) = footnoteFonts(fk) + 1
            Else
                footnoteFonts.Add fk, 1
            End If
        End If
NextFootnote:
    Next fn

    ' ==========================================================
    '  PASS 2: Determine dominant fonts per type
    ' ==========================================================
    Dim domHeading As String
    Dim domBody As String
    Dim domBQ As String
    Dim domFootnote As String

    domHeading = GetDominant(headingFonts)
    domBody = GetDominant(bodyFonts)
    domBQ = GetDominant(bqFonts)
    domFootnote = GetDominant(footnoteFonts)

    ' Only check block_quote type if there are at least 2 paragraphs
    ' (too small a sample otherwise)
    Dim bqTotalCount As Long
    bqTotalCount = 0
    Dim bqK As Variant
    For Each bqK In bqFonts.keys
        bqTotalCount = bqTotalCount + bqFonts(bqK)
    Next bqK
    If bqTotalCount < 2 Then domBQ = ""

    ' ==========================================================
    '  PASS 3: Flag deviations using cached data.
    ' ==========================================================
    Dim paraFontName As String
    Dim paraFontSize As Single

    For pi = 0 To totalParas - 1
        If Not pInRange(pi) Then GoTo NextParaFont2
        If Len(pTypes(pi)) = 0 Then GoTo NextParaFont2

        Dim expectedFont As String
        Dim context As String
        expectedFont = ""
        context = ""

        Select Case pTypes(pi)
            Case "heading"
                If Len(domHeading) > 0 Then
                    expectedFont = domHeading
                    context = "heading"
                End If
            Case "body"
                If Len(domBody) > 0 Then
                    expectedFont = domBody
                    context = "body"
                End If
            Case "block_quote"
                If Len(domBQ) > 0 Then
                    expectedFont = domBQ
                    context = "block quote"
                End If
        End Select

        If Len(expectedFont) = 0 Then GoTo NextParaFont2

        ' -- Check at paragraph level using cached data ----
        paraFontName = pFontNames(pi)
        paraFontSize = pFontSizes(pi)

        If Len(paraFontName) > 0 And paraFontSize > 0 Then
            fk = FontKey(paraFontName, paraFontSize)
            If fk <> expectedFont Then
                Dim findingPara As Object
                Dim locP As String
                Dim paraRng As Range
                Set paraRng = doc.Range(pStarts(pi), pEnds(pi))
                locP = EngineGetLocationString(paraRng, doc)

                Dim cleanParaText As String
                cleanParaText = Trim$(Replace(Left$(paraRng.Text, 60), vbCr, ""))

                Set findingPara = CreateIssueDict(RULE_NAME_FONT, locP, _
                    "Font inconsistency in " & context & ": '" & cleanParaText & _
                    "...' uses " & FontDescription(fk) & " but dominant " & _
                    context & " font is " & FontDescription(expectedFont), _
                    "Change to " & FontDescription(expectedFont), _
                    pStarts(pi), pEnds(pi), "error")
                issues.Add findingPara
                GoTo NextParaFont2
            End If
        End If

        ' -- Run-level check only for mixed-font paragraphs --
        ' (Font info was 0/empty = mixed formatting detected in scan)
        If Len(paraFontName) = 0 Or paraFontSize <= 0 Then
            If pEnds(pi) - pStarts(pi) > 1 Then
                Dim runRange As Range
                Dim runText As String
                Dim isField As Boolean

                Set runRange = doc.Range(pStarts(pi), pEnds(pi))
                runRange.Collapse wdCollapseStart

                On Error Resume Next
                Do While runRange.Start < pEnds(pi)
                    runRange.MoveEnd wdCharacterFormatting, 1
                    If runRange.Start >= pEnds(pi) Then Exit Do

                    Err.Clear
                    runText = runRange.Text
                    If Err.Number <> 0 Then Err.Clear: GoTo AdvanceFontRun

                    If Len(Trim$(Replace(Replace(runText, vbCr, ""), vbLf, ""))) = 0 Then
                        GoTo AdvanceFontRun
                    End If

                    isField = False
                    If runRange.Fields.Count > 0 Then isField = True
                    If Err.Number <> 0 Then Err.Clear: isField = False

                    If Not isField Then
                        fk = FontKey(runRange.Font.Name, runRange.Font.Size)
                        If fk <> expectedFont And Len(runRange.Font.Name) > 0 And _
                           runRange.Font.Size > 0 And runRange.Font.Size < 1000 Then
                            Dim findingRun As Object
                            Dim locR As String
                            Dim cleanRunText As String
                            locR = EngineGetLocationString(runRange, doc)
                            cleanRunText = Trim$(Replace(Left$(runText, 40), vbCr, ""))

                            Set findingRun = CreateIssueDict(RULE_NAME_FONT, locR, _
                                "Mid-paragraph font change in " & context & ": '" & cleanRunText & _
                                "' uses " & FontDescription(fk) & " instead of " & FontDescription(expectedFont), _
                                "Change to " & FontDescription(expectedFont), _
                                runRange.Start, runRange.End, "error")
                            issues.Add findingRun
                            On Error GoTo 0
                            GoTo NextParaFont2
                        End If
                    End If

AdvanceFontRun:
                    runRange.Collapse wdCollapseEnd
                Loop
                On Error GoTo 0
            End If
        End If

NextParaFont2:
    Next pi

    ' ==========================================================
    '  PASS 4: Check footnote font deviations
    ' ==========================================================
    If Len(domFootnote) > 0 Then
        For Each fn In doc.Footnotes
            If Not EngineIsInPageRange(fn.Range) Then GoTo NextFN2

            fnFontName = fn.Range.Font.Name
            fnFontSize = fn.Range.Font.Size

            If Len(fnFontName) > 0 And fnFontSize > 0 And fnFontSize < 1000 Then
                fk = FontKey(fnFontName, fnFontSize)
                If fk <> domFootnote Then
                    Dim findingFN As Object
                    Dim locFN As String
                    locFN = EngineGetLocationString(fn.Range, doc)

                    Dim cleanFNText As String
                    cleanFNText = Trim$(Replace(Left$(fn.Range.Text, 50), vbCr, ""))

                    Set findingFN = CreateIssueDict(RULE_NAME_FONT, locFN, _
                        "Footnote font inconsistency: '" & cleanFNText & _
                        "...' uses " & FontDescription(fk) & " but dominant " & _
                        "footnote font is " & FontDescription(domFootnote), _
                        "Change to " & FontDescription(domFootnote), _
                        fn.Range.Start, fn.Range.End, "error")
                    issues.Add findingFN
                End If
            End If
NextFN2:
        Next fn
    End If

    If Err.Number <> 0 Then
        Debug.Print "Check_FontConsistency: exiting with Err " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    Set Check_FontConsistency = issues
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_Headings.bas

```vb
Attribute VB_Name = "Rules_Headings"
' ============================================================
' Rules_Headings.bas
' Combined module for heading / title rules:
'   - Rule 04: Heading capitalisation consistency
'   - Rule 21: Title (honorific) formatting consistency
'
' Rule 04 uses LOCAL heading families rather than one global
' dominant per outline level.  Headings are grouped into
' contiguous runs separated by structural boundaries
' (appendix/schedule/annex headings, or large gaps of body
' text).  Each family is judged independently.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME_CAPITALISATION As String = "heading_capitalisation"
Private Const RULE_NAME_TITLE As String = "title_formatting"

' Maximum body-text paragraphs between headings before we treat
' the next heading as a new structural family.
Private Const MAX_GAP_PARAS As Long = 40

' --------------------------------------------------------------
'  PRIVATE HELPERS  (from Rule04 - heading capitalisation)
' --------------------------------------------------------------

' -- Minor words to skip when checking Title Case ------------
Private Function GetMinorWords() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim w As Variant
    For Each w In Array("the", "a", "an", "in", "on", "at", "to", _
                        "for", "of", "and", "but", "or", "nor", _
                        "with", "by")
        d.Add CStr(w), True
    Next w
    Set GetMinorWords = d
End Function

' -- Proper nouns that are always capitalised ----------------
Private Function GetProperNouns() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim w As Variant
    For Each w In Array("Court", "Claimant", "Defendant", "Respondent", _
                        "Applicant", "Tribunal", "Parliament", "Crown", _
                        "State", "Government", "Minister")
        d.Add CStr(w), True
    Next w
    Set GetProperNouns = d
End Function

' -- Classify a heading's capitalisation pattern -------------
' Returns "ALL_CAPS", "TITLE_CASE", "SENTENCE_CASE", or "MIXED"
Private Function ClassifyCapitalisation(ByVal headingText As String) As String
    Dim cleanText As String
    Dim i As Long
    Dim ch As String
    Dim hasLower As Boolean
    Dim hasUpper As Boolean

    ' Strip trailing paragraph mark and whitespace
    cleanText = Trim$(Replace(headingText, vbCr, ""))
    cleanText = Trim$(Replace(cleanText, vbLf, ""))
    If Len(cleanText) = 0 Then
        ClassifyCapitalisation = "MIXED"
        Exit Function
    End If

    ' Check ALL CAPS: every alpha character is uppercase
    hasLower = False
    For i = 1 To Len(cleanText)
        ch = Mid$(cleanText, i, 1)
        If ch Like "[a-z]" Then
            hasLower = True
            Exit For
        End If
    Next i
    If Not hasLower Then
        ' Verify there is at least one alpha character
        hasUpper = False
        For i = 1 To Len(cleanText)
            ch = Mid$(cleanText, i, 1)
            If ch Like "[A-Z]" Then
                hasUpper = True
                Exit For
            End If
        Next i
        If hasUpper Then
            ClassifyCapitalisation = "ALL_CAPS"
            Exit Function
        End If
    End If

    ' Split into words and analyse
    Dim words() As String
    words = Split(cleanText, " ")

    Dim minorWords As Object
    Set minorWords = GetMinorWords()

    Dim properNouns As Object
    Set properNouns = GetProperNouns()

    ' Check Title Case: significant words start with uppercase
    Dim titleCaseCount As Long
    Dim significantCount As Long
    Dim wordIdx As Long

    For wordIdx = LBound(words) To UBound(words)
        Dim w As String
        w = Trim$(words(wordIdx))
        If Len(w) = 0 Then GoTo NextWordTitle

        ' Strip leading punctuation
        Dim firstAlpha As String
        Dim charPos As Long
        firstAlpha = ""
        For charPos = 1 To Len(w)
            If Mid$(w, charPos, 1) Like "[A-Za-z]" Then
                firstAlpha = Mid$(w, charPos, 1)
                Exit For
            End If
        Next charPos
        If Len(firstAlpha) = 0 Then GoTo NextWordTitle

        ' Skip minor words (except first word)
        If wordIdx > LBound(words) And minorWords.Exists(LCase(w)) Then
            GoTo NextWordTitle
        End If

        ' Skip proper nouns (always capitalised, not diagnostic)
        If properNouns.Exists(w) Then GoTo NextWordTitle

        significantCount = significantCount + 1
        If firstAlpha Like "[A-Z]" Then
            titleCaseCount = titleCaseCount + 1
        End If
NextWordTitle:
    Next wordIdx

    ' Check Sentence Case: only first word capitalised
    ' First word must start uppercase, rest should be lowercase (except proper nouns)
    Dim firstWord As String
    firstWord = ""
    For wordIdx = LBound(words) To UBound(words)
        If Len(Trim$(words(wordIdx))) > 0 Then
            firstWord = Trim$(words(wordIdx))
            Exit For
        End If
    Next wordIdx

    Dim firstCharOfFirst As String
    firstCharOfFirst = ""
    For charPos = 1 To Len(firstWord)
        If Mid$(firstWord, charPos, 1) Like "[A-Za-z]" Then
            firstCharOfFirst = Mid$(firstWord, charPos, 1)
            Exit For
        End If
    Next charPos

    Dim sentenceCaseViolations As Long
    sentenceCaseViolations = 0
    If firstCharOfFirst Like "[a-z]" Then
        ' First word not capitalised -- not sentence case
        sentenceCaseViolations = significantCount ' force fail
    Else
        ' Check that subsequent significant words start lowercase
        Dim pastFirst As Boolean
        pastFirst = False
        For wordIdx = LBound(words) To UBound(words)
            w = Trim$(words(wordIdx))
            If Len(w) = 0 Then GoTo NextWordSentence
            If Not pastFirst Then
                pastFirst = True
                GoTo NextWordSentence
            End If
            ' Skip proper nouns
            If properNouns.Exists(w) Then GoTo NextWordSentence

            firstAlpha = ""
            For charPos = 1 To Len(w)
                If Mid$(w, charPos, 1) Like "[A-Za-z]" Then
                    firstAlpha = Mid$(w, charPos, 1)
                    Exit For
                End If
            Next charPos
            If Len(firstAlpha) > 0 Then
                If firstAlpha Like "[A-Z]" Then
                    sentenceCaseViolations = sentenceCaseViolations + 1
                End If
            End If
NextWordSentence:
        Next wordIdx
    End If

    ' Determine pattern
    If significantCount > 0 And titleCaseCount = significantCount Then
        ClassifyCapitalisation = "TITLE_CASE"
    ElseIf sentenceCaseViolations = 0 Then
        ClassifyCapitalisation = "SENTENCE_CASE"
    Else
        ClassifyCapitalisation = "MIXED"
    End If
End Function

' -- Count words in a heading (excluding trailing marks) -----
Private Function CountWords(ByVal txt As String) As Long
    Dim cleanText As String
    cleanText = Trim$(Replace(txt, vbCr, ""))
    cleanText = Trim$(Replace(cleanText, vbLf, ""))
    If Len(cleanText) = 0 Then
        CountWords = 0
        Exit Function
    End If
    Dim wParts() As String
    wParts = Split(cleanText, " ")
    Dim cnt As Long
    Dim p As Variant
    For Each p In wParts
        If Len(Trim$(CStr(p))) > 0 Then cnt = cnt + 1
    Next p
    CountWords = cnt
End Function

' -- Check if a heading text indicates a structural boundary --
'  Returns True for schedule, appendix, annex, part, exhibit etc.
Private Function IsStructuralBoundary(ByVal headingText As String) As Boolean
    Dim lText As String
    lText = LCase$(Trim$(Replace(headingText, vbCr, "")))
    IsStructuralBoundary = False

    ' Check for section-divider keywords at the start
    If Left$(lText, 8) = "schedule" Or Left$(lText, 8) = "appendix" Or _
       Left$(lText, 5) = "annex" Or Left$(lText, 7) = "exhibit" Or _
       Left$(lText, 10) = "attachment" Then
        IsStructuralBoundary = True
        Exit Function
    End If

    ' Also match "SCHEDULE", "APPENDIX" etc. in ALL_CAPS
    If Left$(lText, 4) = "part" Then
        ' "Part" followed by a number or letter is structural
        If Len(lText) > 4 Then
            Dim afterPart As String
            afterPart = Trim$(Mid$(lText, 5))
            If Len(afterPart) > 0 Then
                Dim fc As String
                fc = Left$(afterPart, 1)
                If (fc >= "0" And fc <= "9") Or _
                   (fc >= "a" And fc <= "z") Then
                    IsStructuralBoundary = True
                End If
            End If
        End If
    End If
End Function

' --------------------------------------------------------------
'  PRIVATE HELPERS  (from Rule21 - title formatting)
' --------------------------------------------------------------

' -- Count occurrences of a word in the document -------------
'  Uses Find with MatchWholeWord and MatchCase.
Private Function CountWordInDoc(doc As Document, word As String) As Long
    Dim rng As Range
    Dim cnt As Long
    Dim found As Boolean

    cnt = 0

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = word
        .MatchWholeWord = True
        .MatchCase = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do

        If EngineIsInPageRange(rng) Then
            cnt = cnt + 1
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    CountWordInDoc = cnt
End Function

' -- Flag all occurrences of a minority form -----------------
Private Sub FlagOccurrences(doc As Document, _
                             word As String, _
                             issueText As String, _
                             suggestionText As String, _
                             ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = word
        .MatchWholeWord = True
        .MatchCase = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do

        If EngineIsInPageRange(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_TITLE, locStr, issueText, suggestionText, rng.Start, rng.End, "error")
            issues.Add finding
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ============================================================
'  PUBLIC: Check heading capitalisation  (Rule 04)
'
'  LOCAL-FAMILY APPROACH:
'  1. Collect all headings into an ordered list.
'  2. Walk the ordered list and split into "families" whenever:
'     a) A structural boundary heading is encountered
'        (schedule, appendix, annex, etc.), or
'     b) More than MAX_GAP_PARAS non-heading paragraphs
'        separate two consecutive headings.
'  3. Within each family, determine dominant capitalisation
'     per outline level and flag outliers.
' ============================================================
Public Function Check_HeadingCapitalisation(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim lvl As Long

    On Error Resume Next

    ' -------------------------------------------------------
    '  PASS 1: Collect all headings into an ordered array
    ' -------------------------------------------------------
    ' Each entry: Array(paraIdx, headingText, pattern, rangeStart, rangeEnd, outlineLevel)
    Dim hCap As Long
    hCap = 128
    Dim hCount As Long
    hCount = 0
    Dim hParaIdx() As Long
    Dim hTexts() As String
    Dim hPatterns() As String
    Dim hStarts() As Long
    Dim hEnds() As Long
    Dim hLevels() As Long
    ReDim hParaIdx(0 To hCap - 1)
    ReDim hTexts(0 To hCap - 1)
    ReDim hPatterns(0 To hCap - 1)
    ReDim hStarts(0 To hCap - 1)
    ReDim hEnds(0 To hCap - 1)
    ReDim hLevels(0 To hCap - 1)

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        ' Check if this is a heading (outline levels 1-9)
        lvl = para.OutlineLevel
        If Err.Number <> 0 Then lvl = wdOutlineLevelBodyText: Err.Clear
        If lvl >= wdOutlineLevel1 And lvl <= wdOutlineLevel9 Then

            ' Page range filter
            If Not EngineIsInPageRange(para.Range) Then GoTo NextPara

            Dim headingText As String
            headingText = para.Range.Text
            If Err.Number <> 0 Then headingText = "": Err.Clear

            ' Skip single-word headings
            If CountWords(headingText) <= 1 Then GoTo NextPara

            ' Classify capitalisation
            Dim pattern As String
            pattern = ClassifyCapitalisation(headingText)

            ' Grow arrays if needed
            If hCount >= hCap Then
                hCap = hCap * 2
                ReDim Preserve hParaIdx(0 To hCap - 1)
                ReDim Preserve hTexts(0 To hCap - 1)
                ReDim Preserve hPatterns(0 To hCap - 1)
                ReDim Preserve hStarts(0 To hCap - 1)
                ReDim Preserve hEnds(0 To hCap - 1)
                ReDim Preserve hLevels(0 To hCap - 1)
            End If

            hParaIdx(hCount) = paraIdx
            hTexts(hCount) = headingText
            hPatterns(hCount) = pattern
            hStarts(hCount) = para.Range.Start
            If Err.Number <> 0 Then hStarts(hCount) = 0: Err.Clear
            hEnds(hCount) = para.Range.End
            If Err.Number <> 0 Then hEnds(hCount) = 0: Err.Clear
            hLevels(hCount) = lvl
            hCount = hCount + 1
        End If
NextPara:
    Next para

    If hCount < 2 Then
        On Error GoTo 0
        Set Check_HeadingCapitalisation = issues
        Exit Function
    End If
    On Error GoTo 0   ' Pass 1 complete; Pass 2 is pure VBA

    ' -------------------------------------------------------
    '  PASS 2: Split headings into local families
    '
    '  familyStarts() and familyEnds() mark index ranges
    '  within the heading arrays.
    ' -------------------------------------------------------
    Dim fsCap As Long
    fsCap = 32
    Dim fsCount As Long
    fsCount = 0
    Dim familyStarts() As Long
    Dim familyEnds() As Long
    ReDim familyStarts(0 To fsCap - 1)
    ReDim familyEnds(0 To fsCap - 1)

    Dim curFamilyStart As Long
    curFamilyStart = 0

    Dim hi As Long
    For hi = 1 To hCount - 1
        Dim newFamily As Boolean
        newFamily = False

        ' Check for structural boundary
        If IsStructuralBoundary(hTexts(hi)) Then
            newFamily = True
        End If

        ' Check for large gap between consecutive headings
        If Not newFamily Then
            Dim gap As Long
            gap = hParaIdx(hi) - hParaIdx(hi - 1)
            If gap > MAX_GAP_PARAS Then
                newFamily = True
            End If
        End If

        If newFamily Then
            ' Close the current family
            If fsCount >= fsCap Then
                fsCap = fsCap * 2
                ReDim Preserve familyStarts(0 To fsCap - 1)
                ReDim Preserve familyEnds(0 To fsCap - 1)
            End If
            familyStarts(fsCount) = curFamilyStart
            familyEnds(fsCount) = hi - 1
            fsCount = fsCount + 1
            curFamilyStart = hi
        End If
    Next hi

    ' Close the last family
    If fsCount >= fsCap Then
        fsCap = fsCap * 2
        ReDim Preserve familyStarts(0 To fsCap - 1)
        ReDim Preserve familyEnds(0 To fsCap - 1)
    End If
    familyStarts(fsCount) = curFamilyStart
    familyEnds(fsCount) = hCount - 1
    fsCount = fsCount + 1

    ' -------------------------------------------------------
    '  PASS 3: Within each family, find dominant per level
    '  and flag outliers
    ' -------------------------------------------------------
    Dim fi As Long
    For fi = 0 To fsCount - 1
        Dim fStart As Long
        fStart = familyStarts(fi)
        Dim fEnd As Long
        fEnd = familyEnds(fi)

        ' Build pattern counts per level within this family
        Dim levelPats As Object
        Set levelPats = CreateObject("Scripting.Dictionary")
        ' levelPats: level -> Dictionary(pattern -> count)

        Dim hj As Long
        For hj = fStart To fEnd
            lvl = hLevels(hj)
            If Not levelPats.Exists(lvl) Then
                levelPats.Add lvl, CreateObject("Scripting.Dictionary")
            End If
            Dim patDict As Object
            Set patDict = levelPats(lvl)
            If patDict.Exists(hPatterns(hj)) Then
                patDict(hPatterns(hj)) = patDict(hPatterns(hj)) + 1
            Else
                patDict.Add hPatterns(hj), 1
            End If
        Next hj

        ' For each level in this family, find dominant and flag outliers
        Dim lvlKey As Variant
        For Each lvlKey In levelPats.keys
            Set patDict = levelPats(lvlKey)

            ' Count total headings at this level in this family
            Dim levelTotal As Long
            levelTotal = 0
            Dim patKey As Variant
            For Each patKey In patDict.keys
                levelTotal = levelTotal + patDict(patKey)
            Next patKey

            ' Need at least 2 headings to compare
            If levelTotal < 2 Then GoTo NextFamilyLevel

            ' Find dominant pattern
            Dim dominantPattern As String
            Dim maxCount As Long
            dominantPattern = ""
            maxCount = 0
            For Each patKey In patDict.keys
                If patDict(patKey) > maxCount Then
                    maxCount = patDict(patKey)
                    dominantPattern = CStr(patKey)
                End If
            Next patKey

            ' Flag headings in this family+level that deviate
            For hj = fStart To fEnd
                If hLevels(hj) = CLng(lvlKey) Then
                    If hPatterns(hj) <> dominantPattern Then
                        Dim finding As Object
                        Dim loc As String
                        Dim rng As Range
                        On Error Resume Next
                        Set rng = doc.Range(hStarts(hj), hEnds(hj))
                        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextFamilyHeading
                        loc = EngineGetLocationString(rng, doc)
                        If Err.Number <> 0 Then loc = "unknown location": Err.Clear
                        On Error GoTo 0

                        Dim cleanHText As String
                        cleanHText = Trim$(Replace(hTexts(hj), vbCr, ""))

                        Dim suggn As String
                        Select Case dominantPattern
                            Case "ALL_CAPS"
                                suggn = "Convert to ALL CAPS to match nearby level " & CLng(lvlKey) & " headings"
                            Case "TITLE_CASE"
                                suggn = "Convert to Title Case to match nearby level " & CLng(lvlKey) & " headings"
                            Case "SENTENCE_CASE"
                                suggn = "Convert to Sentence case to match nearby level " & CLng(lvlKey) & " headings"
                            Case Else
                                suggn = "Review capitalisation for consistency with nearby level " & CLng(lvlKey) & " headings"
                        End Select

                        Set finding = CreateIssueDict(RULE_NAME_CAPITALISATION, loc, "Heading capitalisation mismatch: '" & cleanHText & "' uses " & hPatterns(hj) & " but nearby dominant pattern is " & dominantPattern, suggn, hStarts(hj), hEnds(hj), "possible_error")
                        issues.Add finding
                    End If
                End If
NextFamilyHeading:
            Next hj
NextFamilyLevel:
        Next lvlKey
    Next fi

    On Error GoTo 0
    Set Check_HeadingCapitalisation = issues
End Function

' ============================================================
'  PUBLIC: Check title formatting  (Rule 21)
' ============================================================
Public Function Check_TitleFormatting(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Define title pairs: noDot / withDot ------------------
    Dim noDot As Variant
    Dim withDot As Variant

    noDot = Array("Mr", "Mrs", "Ms", "Dr", "Prof", "QC", "KC", "MP", "JP")
    withDot = Array("Mr.", "Mrs.", "Ms.", "Dr.", "Prof.", "Q.C.", "K.C.", "M.P.", "J.P.")

    Dim i As Long
    Dim noDotCount As Long
    Dim withDotCount As Long

    For i = LBound(noDot) To UBound(noDot)
        noDotCount = CountWordInDoc(doc, CStr(noDot(i)))
        withDotCount = CountWordInDoc(doc, CStr(withDot(i)))

        ' Only flag if both forms exist
        If noDotCount > 0 And withDotCount > 0 Then
            If noDotCount >= withDotCount Then
                ' noDot is dominant -- flag all withDot occurrences
                FlagOccurrences doc, CStr(withDot(i)), _
                    "Inconsistent title formatting: '" & withDot(i) & "' used", _
                    "Use '" & noDot(i) & "' without full stop (dominant style)", _
                    issues
            Else
                ' withDot is dominant -- flag all noDot occurrences
                FlagOccurrences doc, CStr(noDot(i)), _
                    "Inconsistent title formatting: '" & noDot(i) & "' used", _
                    "Use '" & withDot(i) & "' with full stop (dominant style)", _
                    issues
            End If
        End If
    Next i

    Set Check_TitleFormatting = issues
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_Italics.bas

```vb
Attribute VB_Name = "Rules_Italics"
' ============================================================
' Rules_Italics.bas
' Combined italics-related proofreading rules:
'   - Rule 30: flags italicisation of known anglicised foreign
'     terms that should be set in roman (upright) type.
'   - Rule 31: flags italicisation of foreign names, institutions,
'     places or courts that should not be italicised.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

' -- Rule-name constants -------------------------------------
Private Const RULE_NAME_ANGLICISED As String = "known_anglicised_terms_not_italic"
Private Const RULE_NAME_FOREIGN   As String = "foreign_names_not_italic"

' -- Seed list of anglicised terms (Rule 30) -----------------
Private seedTerms As Variant
Private seedInitialised As Boolean

' -- Module-level dictionary of protected foreign names (Rule 31) -
' Key = name (String), Value = True (Boolean) -- used as a set
Private foreignNames As Object

' ============================================================
'  SHARED PRIVATE HELPERS
' ============================================================

' ------------------------------------------------------------
'  Check whether a character is a letter (A-Z, a-z)
' ------------------------------------------------------------
Private Function IsLetter(ByVal ch As String) As Boolean
    Dim c As Long
    If Len(ch) = 0 Then
        IsLetter = False
        Exit Function
    End If
    c = AscW(Left$(ch, 1))
    IsLetter = (c >= 65 And c <= 90) Or (c >= 97 And c <= 122)
End Function

' ------------------------------------------------------------
'  Check whether a range span is italic
'  Returns True if any part of the range is italic.
' ------------------------------------------------------------
Private Function IsRangeItalic(rng As Range) As Boolean
    Dim italVal As Long

    On Error Resume Next
    italVal = rng.Font.Italic
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: IsRangeItalic = False: Exit Function
    On Error GoTo 0

    ' If Font.Italic is True the whole range is italic
    If italVal = True Then
        IsRangeItalic = True
        Exit Function
    End If

    ' wdToggle treated as italic present
    If italVal = wdToggle Then
        IsRangeItalic = True
        Exit Function
    End If

    ' If Font.Italic is wdUndefined (9999999) the range has
    ' mixed formatting -- check individual characters
    If italVal = wdUndefined Then
        Dim i As Long
        Dim charRng As Range
        For i = rng.Start To rng.End - 1
            On Error Resume Next
            Set charRng = rng.Document.Range(i, i + 1)
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextCharItalic
            Dim charItal As Long
            charItal = charRng.Font.Italic
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextCharItalic
            On Error GoTo 0
            If charItal = True Then
                IsRangeItalic = True
                Exit Function
            End If
NextCharItalic:
        Next i
    End If

    IsRangeItalic = False
End Function

' ============================================================
'  RULE 30 -- ANGLICISED TERMS NOT ITALIC
' ============================================================

' ------------------------------------------------------------
'  Initialise the seed term list
' ------------------------------------------------------------
Private Sub InitSeedTerms()
    If seedInitialised Then Exit Sub
    Dim batch1 As Variant, batch2 As Variant, batch3 As Variant
    batch1 = Array( _
        "amicus curiae", "a priori", "a fortiori", "bona fide", _
        "de facto", "de jure", "ex parte", "ex post", _
        "ex post facto", "indicia")
    batch2 = Array( _
        "in situ", "inter alia", "laissez-faire", "mutatis mutandis", _
        "novus actus interveniens", "obiter dicta", "per se", _
        "prima facie", "quantum meruit", "quid pro quo")
    batch3 = Array( _
        "raison d'etre", "ratio decidendi", "stare decisis", _
        "terra nullius", "ultra vires", "vice versa", _
        "vis-a-vis", "viz")
    seedTerms = MergeArrays(batch1, batch2, batch3)
    seedInitialised = True
End Sub

' ------------------------------------------------------------
'  MAIN ENTRY POINT -- Rule 30
' ------------------------------------------------------------
Public Function Check_AnglicisedTermsNotItalic(doc As Document) As Collection
    Dim issues As New Collection

    InitSeedTerms

    Dim para As Paragraph
    Dim paraText As String
    Dim pos As Long
    Dim termIdx As Long
    Dim term As String
    Dim termLen As Long
    Dim charBefore As String
    Dim charAfter As String
    Dim rng As Range
    Dim locStr As String
    Dim finding As Object

    For Each para In doc.Paragraphs
        ' Skip paragraphs outside the configured page range
        On Error Resume Next
        Dim inRange As Boolean
        inRange = EngineIsInPageRange(para.Range)
        If Err.Number <> 0 Then inRange = True: Err.Clear
        On Error GoTo 0
        If Not inRange Then GoTo NextParaR30

        On Error Resume Next
        paraText = para.Range.Text
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextParaR30
        On Error GoTo 0
        If Len(paraText) = 0 Then GoTo NextParaR30

        For termIdx = LBound(seedTerms) To UBound(seedTerms)
            term = CStr(seedTerms(termIdx))
            termLen = Len(term)

            pos = InStr(1, paraText, term, vbTextCompare)
            Do While pos > 0
                If pos > 1 Then
                    charBefore = Mid$(paraText, pos - 1, 1)
                    If IsLetter(charBefore) Then GoTo NextMatchR30
                End If

                If pos + termLen <= Len(paraText) Then
                    charAfter = Mid$(paraText, pos + termLen, 1)
                    If IsLetter(charAfter) Then GoTo NextMatchR30
                End If

                On Error Resume Next
                Set rng = doc.Range( _
                    para.Range.Start + pos - 1, _
                    para.Range.Start + pos - 1 + termLen)
                If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextMatchR30
                On Error GoTo 0

                If IsRangeItalic(rng) Then
                    On Error Resume Next
                    locStr = EngineGetLocationString(rng, doc)
                    If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
                    On Error GoTo 0

                    Set finding = CreateIssueDict(RULE_NAME_ANGLICISED, locStr, "Anglicised foreign term is italicised.", "Set '" & term & "' in roman, not italics.", rng.Start, rng.End, "warning", False)
                    issues.Add finding
                End If

NextMatchR30:
                pos = InStr(pos + 1, paraText, term, vbTextCompare)
            Loop
        Next termIdx

NextParaR30:
    Next para

    Set Check_AnglicisedTermsNotItalic = issues
End Function

' ============================================================
'  RULE 31 -- FOREIGN NAMES NOT ITALIC
' ============================================================

' ------------------------------------------------------------
'  Initialise the seed name dictionary
' ------------------------------------------------------------
Private Sub InitSeedNames()
    Set foreignNames = CreateObject("Scripting.Dictionary")
    foreignNames.CompareMode = vbTextCompare

    foreignNames.Add "Cour de cassation", True
    foreignNames.Add "Conseil d'Etat", True
    foreignNames.Add "Bayerisches Staatsministerium der Justiz", True
End Sub

' ------------------------------------------------------------
'  PUBLIC: Add a foreign name to the protected list
' ------------------------------------------------------------
Public Sub AddForeignName(ByVal termName As String)
    If foreignNames Is Nothing Then
        InitSeedNames
    End If

    If Not foreignNames.Exists(termName) Then
        foreignNames.Add termName, True
    End If
End Sub

' ------------------------------------------------------------
'  MAIN ENTRY POINT -- Rule 31
' ------------------------------------------------------------
Public Function Check_ForeignNamesNotItalic(doc As Document) As Collection
    Dim issues As New Collection

    ' Initialise defaults if not yet loaded
    If foreignNames Is Nothing Then
        InitSeedNames
    End If

    Dim para As Paragraph
    Dim paraText As String
    Dim pos As Long
    Dim nameKey As Variant
    Dim term As String
    Dim termLen As Long
    Dim charBefore As String
    Dim charAfter As String
    Dim rng As Range
    Dim locStr As String
    Dim finding As Object
    Dim keys As Variant

    keys = foreignNames.keys

    For Each para In doc.Paragraphs
        On Error Resume Next
        Dim inRange31 As Boolean
        inRange31 = EngineIsInPageRange(para.Range)
        If Err.Number <> 0 Then inRange31 = True: Err.Clear
        On Error GoTo 0
        If Not inRange31 Then GoTo NextParaR31

        On Error Resume Next
        paraText = para.Range.Text
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextParaR31
        On Error GoTo 0
        If Len(paraText) = 0 Then GoTo NextParaR31

        Dim k As Long
        For k = 0 To foreignNames.Count - 1
            term = CStr(keys(k))
            termLen = Len(term)

            pos = InStr(1, paraText, term, vbTextCompare)
            Do While pos > 0
                If pos > 1 Then
                    charBefore = Mid$(paraText, pos - 1, 1)
                    If IsLetter(charBefore) Then GoTo NextMatchR31
                End If

                If pos + termLen <= Len(paraText) Then
                    charAfter = Mid$(paraText, pos + termLen, 1)
                    If IsLetter(charAfter) Then GoTo NextMatchR31
                End If

                On Error Resume Next
                Set rng = doc.Range( _
                    para.Range.Start + pos - 1, _
                    para.Range.Start + pos - 1 + termLen)
                If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextMatchR31
                On Error GoTo 0

                If IsRangeItalic(rng) Then
                    On Error Resume Next
                    locStr = EngineGetLocationString(rng, doc)
                    If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
                    On Error GoTo 0

                    Set finding = CreateIssueDict(RULE_NAME_FOREIGN, locStr, "Foreign name or institution should not be italicised.", "Set '" & term & "' in roman, not italics.", rng.Start, rng.End, "warning", False)
                    issues.Add finding
                End If

NextMatchR31:
                pos = InStr(pos + 1, paraText, term, vbTextCompare)
            Loop
        Next k

NextParaR31:
    Next para

    Set Check_ForeignNamesNotItalic = issues
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Merge up to 3 Variant arrays into one flat Variant array
' ----------------------------------------------------------------
Private Function MergeArrays(a1 As Variant, a2 As Variant, a3 As Variant) As Variant
    Dim total As Long
    total = UBound(a1) - LBound(a1) + 1 _
          + UBound(a2) - LBound(a2) + 1 _
          + UBound(a3) - LBound(a3) + 1
    Dim out() As Variant
    ReDim out(0 To total - 1)
    Dim idx As Long
    idx = 0
    Dim v As Variant
    For Each v In a1: out(idx) = v: idx = idx + 1: Next v
    For Each v In a2: out(idx) = v: idx = idx + 1: Next v
    For Each v In a3: out(idx) = v: idx = idx + 1: Next v
    MergeArrays = out
End Function

```

# FILE: Rules_LegalTerms.bas

```vb
Attribute VB_Name = "Rules_LegalTerms"
' ============================================================
' Rules_LegalTerms.bas
' Combined module for Rule28 (mandated legal term forms) and
' Rule29 (always capitalise terms).
'
' Rule28: enforces fixed hyphenation for specific legal and
'   governmental terms. Flags unhyphenated variants and suggests
'   the approved hyphenated form.
'   Default mandatory list:
'     "Solicitor-General", "Attorney-General"
'   Additional terms can be added at runtime via AddMandatedTerm.
'
' Rule29: enforces capitalisation for specified Hart-style terms.
'   Scans each paragraph for case-insensitive matches and flags
'   any occurrence whose capitalisation does not match the
'   approved form. Matches inside quoted material are skipped.
'   Context-sensitive terms (Province, State, party names) are
'   intentionally omitted -- the engine does not yet have
'   reliable context handling.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE28_NAME As String = "mandated_legal_term_forms"
Private Const RULE29_NAME As String = "always_capitalise_terms"

' -- Module-level dictionary for Rule28 ----------------------
' Key = LCase(correct form), Value = correct form (String)
Private mandatedTerms As Object

' ============================================================
'  RULE 28 -- MAIN ENTRY POINT
' ============================================================
Public Function Check_MandatedLegalTermForms(doc As Document) As Collection
    Dim issues As New Collection

    ' Initialise defaults if not yet loaded
    If mandatedTerms Is Nothing Then
        InitDefaultTerms
    End If

    Dim keys As Variant
    Dim k As Long
    Dim correctForm As String
    Dim searchPhrase As String

    keys = mandatedTerms.keys

    For k = 0 To mandatedTerms.Count - 1
        correctForm = CStr(mandatedTerms(keys(k)))

        ' Build the unhyphenated search variant by replacing hyphens with spaces
        searchPhrase = Replace(correctForm, "-", " ")

        ' Only search if the unhyphenated form is actually different
        If StrComp(searchPhrase, correctForm, vbBinaryCompare) <> 0 Then
            SearchAndFlag doc, searchPhrase, correctForm, issues
        End If
    Next k

    Set Check_MandatedLegalTermForms = issues
End Function

' ============================================================
'  RULE 28 -- PRIVATE: Search for an unhyphenated variant and flag matches
' ============================================================
Private Sub SearchAndFlag(doc As Document, _
                           searchPhrase As String, _
                           correctForm As String, _
                           ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = searchPhrase
        .MatchWholeWord = True
        .MatchCase = False
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        ' Skip if the matched text already has the correct hyphenated form
        If StrComp(rng.Text, correctForm, vbTextCompare) = 0 Then
            GoTo SkipMatch
        End If

        ' Verify it is not actually the hyphenated form by checking
        ' the surrounding context -- the Find matched with MatchCase=False
        ' and spaces, so an exact binary comparison rules out false positives
        If EngineIsInPageRange(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE28_NAME, locStr, "Mandatory term is not hyphenated in the approved form.", "Use '" & correctForm & "'.", rng.Start, rng.End, "warning", False)
            issues.Add finding
        End If

SkipMatch:
        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ============================================================
'  RULE 28 -- PRIVATE: Populate default mandatory terms
' ============================================================
Private Sub InitDefaultTerms()
    Set mandatedTerms = CreateObject("Scripting.Dictionary")

    mandatedTerms.Add LCase("Solicitor-General"), "Solicitor-General"
    mandatedTerms.Add LCase("Attorney-General"), "Attorney-General"
End Sub

' ============================================================
'  RULE 28 -- PUBLIC: Add a mandated term at runtime
'  The term must contain a hyphen (e.g. "Director-General").
'  If the term already exists it is silently ignored.
' ============================================================
Public Sub AddMandatedTerm(term As String)
    If mandatedTerms Is Nothing Then
        InitDefaultTerms
    End If

    Dim lcKey As String
    lcKey = LCase(term)

    If Not mandatedTerms.Exists(lcKey) Then
        mandatedTerms.Add lcKey, term
    End If
End Sub

' ============================================================
'  RULE 29 -- MAIN ENTRY POINT
' ============================================================
Public Function Check_AlwaysCapitaliseTerms(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Seed dictionary of correct forms -------------------
    Dim terms As Variant
    Dim batch1 As Variant, batch2 As Variant
    batch1 = Array( _
        "Act", "Bill", "Attorney-General", "Cabinet", _
        "Commonwealth", "Constitution", "Crown", _
        "Executive Council", "Governor", "Governor-General", _
        "Her Majesty", "the Queen")
    batch2 = Array( _
        "his Honour", "her Honour", "their Honours", _
        "Law Lords", "their Lordships", "Lords Justices", _
        "Member States", "Parliament", "Labour Party", _
        "Prime Minister", "Vice-Chancellor")
    terms = MergeArrays2(batch1, batch2)

    ' -- Iterate paragraphs ---------------------------------
    Dim para As Paragraph
    Dim paraRng As Range
    Dim paraText As String
    Dim paraStart As Long

    For Each para In doc.Paragraphs
        On Error Resume Next
        Set paraRng = para.Range
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextPara
        On Error GoTo 0

        ' Check page range filter
        If Not EngineIsInPageRange(paraRng) Then GoTo NextPara

        paraText = paraRng.Text
        paraStart = paraRng.Start

        If Len(paraText) = 0 Then GoTo NextPara

        ' -- Check each term against this paragraph ---------
        Dim t As Long
        For t = LBound(terms) To UBound(terms)
            CheckTermInParagraph doc, CStr(terms(t)), paraText, paraStart, paraRng, issues
        Next t

NextPara:
    Next para

    Set Check_AlwaysCapitaliseTerms = issues
End Function

' ============================================================
'  RULE 29 -- PRIVATE: Search for a single term within one paragraph
' ============================================================
Private Sub CheckTermInParagraph(doc As Document, _
                                  correctForm As String, _
                                  paraText As String, _
                                  paraStart As Long, _
                                  paraRng As Range, _
                                  ByRef issues As Collection)
    Dim termLen As Long
    Dim pos As Long
    Dim actualText As String
    Dim matchStart As Long
    Dim matchEnd As Long
    Dim finding As Object
    Dim locStr As String
    Dim charBefore As String
    Dim charAfter As String

    termLen = Len(correctForm)

    ' Walk through all case-insensitive matches in the paragraph
    pos = InStr(1, paraText, correctForm, vbTextCompare)

    Do While pos > 0
        ' -- Word boundary check ----------------------------
        ' Ensure we are not matching a substring of a longer word
        If pos > 1 Then
            charBefore = Mid(paraText, pos - 1, 1)
            If IsWordChar(charBefore) Then GoTo NextMatch
        End If

        If pos + termLen <= Len(paraText) Then
            charAfter = Mid(paraText, pos + termLen, 1)
            If IsWordChar(charAfter) Then GoTo NextMatch
        End If

        ' -- Extract the actual text at the match position --
        actualText = Mid(paraText, pos, termLen)

        ' -- Skip if capitalisation already matches ---------
        If StrComp(actualText, correctForm, vbBinaryCompare) = 0 Then
            GoTo NextMatch
        End If

        ' -- Skip if inside quoted material -----------------
        If IsInsideQuote(paraText, pos) Then GoTo NextMatch

        ' -- Calculate range positions ----------------------
        matchStart = paraStart + pos - 1
        matchEnd = matchStart + termLen

        On Error Resume Next
        Dim matchRng As Range
        Set matchRng = doc.Range(matchStart, matchEnd)
        locStr = EngineGetLocationString(matchRng, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
        On Error GoTo 0

        Set finding = CreateIssueDict(RULE29_NAME, locStr, "Term should be capitalised in the approved form.", "Use '" & correctForm & "'.", matchStart, matchEnd, "warning", False)
        issues.Add finding

NextMatch:
        ' Search for next occurrence after current position
        If pos + 1 > Len(paraText) Then Exit Do
        pos = InStr(pos + 1, paraText, correctForm, vbTextCompare)
    Loop
End Sub

' ============================================================
'  PRIVATE: Check whether a character is a word character
'  (letter, digit, hyphen, or underscore)
' ============================================================
Private Function IsWordChar(ch As String) As Boolean
    Dim c As Long
    If Len(ch) = 0 Then
        IsWordChar = False
        Exit Function
    End If

    c = AscW(ch)

    ' A-Z, a-z
    If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Then
        IsWordChar = True
        Exit Function
    End If

    ' 0-9
    If c >= 48 And c <= 57 Then
        IsWordChar = True
        Exit Function
    End If

    ' Hyphen or underscore (treat as word chars for compound terms)
    If c = 45 Or c = 95 Then
        IsWordChar = True
        Exit Function
    End If

    IsWordChar = False
End Function

' ============================================================
'  PRIVATE: Determine if position is inside quoted material
'  Checks for smart quotes and straight quotes in a window
'  before the match position.
' ============================================================
Private Function IsInsideQuote(paraText As String, matchPos As Long) As Boolean
    Dim openCount As Long
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim windowStart As Long

    IsInsideQuote = False
    openCount = 0

    ' Scan from start of paragraph to match position
    ' to count unmatched opening quotes
    windowStart = 1
    If matchPos <= 1 Then Exit Function

    For i = windowStart To matchPos - 1
        ch = Mid(paraText, i, 1)
        code = AscW(ch)

        Select Case code
            Case 8220  ' left double smart quote
                openCount = openCount + 1
            Case 8221  ' right double smart quote
                If openCount > 0 Then openCount = openCount - 1
            Case 8216  ' left single smart quote
                ' Skip if flanked by letters (apostrophe in smart-quote mode)
                If i > 1 And i < Len(paraText) Then
                    Dim ls16Prev As String, ls16Next As String
                    ls16Prev = Mid(paraText, i - 1, 1)
                    ls16Next = Mid(paraText, i + 1, 1)
                    Dim ls16PrevA As Boolean, ls16NextA As Boolean
                    ls16PrevA = IsWordChar(ls16Prev) And ls16Prev <> "-" And ls16Prev <> "_"
                    ls16NextA = IsWordChar(ls16Next) And ls16Next <> "-" And ls16Next <> "_"
                    If Not (ls16PrevA And ls16NextA) Then
                        openCount = openCount + 1
                    End If
                Else
                    openCount = openCount + 1
                End If
            Case 8217  ' right single smart quote
                ' Skip if flanked by letters (apostrophe: it's, don't)
                If i > 1 And i < Len(paraText) Then
                    Dim rs17Prev As String, rs17Next As String
                    rs17Prev = Mid(paraText, i - 1, 1)
                    rs17Next = Mid(paraText, i + 1, 1)
                    Dim rs17PrevA As Boolean, rs17NextA As Boolean
                    rs17PrevA = IsWordChar(rs17Prev) And rs17Prev <> "-" And rs17Prev <> "_"
                    rs17NextA = IsWordChar(rs17Next) And rs17Next <> "-" And rs17Next <> "_"
                    If rs17PrevA And rs17NextA Then
                        ' Apostrophe - skip
                    ElseIf openCount > 0 Then
                        openCount = openCount - 1
                    End If
                ElseIf openCount > 0 Then
                    openCount = openCount - 1
                End If
            Case 34    ' straight double quote -- toggle
                If openCount > 0 Then
                    openCount = openCount - 1
                Else
                    openCount = openCount + 1
                End If
            Case 39    ' straight single quote / apostrophe
                ' Distinguish apostrophe from quote delimiter:
                ' If flanked by letters/digits on both sides, it's an apostrophe -> skip.
                ' Otherwise use whitespace heuristic for open/close.
                Dim prevCh As String
                Dim nextCh As String
                Dim prevIsAlpha As Boolean, nextIsAlpha As Boolean
                prevIsAlpha = False
                nextIsAlpha = False
                If i > 1 Then
                    prevCh = Mid(paraText, i - 1, 1)
                    prevIsAlpha = IsWordChar(prevCh) And prevCh <> "-" And prevCh <> "_"
                End If
                If i < Len(paraText) Then
                    nextCh = Mid(paraText, i + 1, 1)
                    nextIsAlpha = IsWordChar(nextCh) And nextCh <> "-" And nextCh <> "_"
                End If
                ' Letter/digit on both sides = apostrophe (it's, don't, 90's)
                If prevIsAlpha And nextIsAlpha Then
                    ' Skip: this is an apostrophe, not a quote
                Else
                    If i = 1 Then
                        openCount = openCount + 1
                    ElseIf Not prevIsAlpha Then
                        ' Preceded by space/punct = opening quote
                        openCount = openCount + 1
                    ElseIf openCount > 0 Then
                        openCount = openCount - 1
                    End If
                End If
        End Select
    Next i

    ' If there are unmatched opening quotes, the match is inside quoted material
    If openCount > 0 Then
        IsInsideQuote = True
    End If
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

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

```

# FILE: Rules_Lists.bas

```vb
Attribute VB_Name = "Rules_Lists"
' ============================================================
' Rules_Lists.bas
' Combined module for list-related proofreading rules:
'   - Rule10: Inline list format consistency (separator style,
'     conjunction usage, ending punctuation)
'   - Rule15: List punctuation consistency (ending punctuation
'     of formal list items, final-item full stop, penultimate
'     conjunction)
'
' ENGINE WIRING NOTE:
'   Both rules are dispatched under the single aggregate toggle
'   "list_rules" in PleadingsEngine.InitRuleConfig / RunAllPleadingsRules.
'   Enabling/disabling "list_rules" controls both Check_InlineListFormat
'   and Check_ListPunctuation together.
'
' Rule 10 uses LOCAL-CONTEXT grouping: only inline lists that
' are structurally close (within the same section-like region)
' are compared for consistency.  Unrelated lists in different
' sections are judged independently.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

' -- Rule name constants ---------------------------------------
Private Const RULE_NAME_INLINE  As String = "inline_list_format"
Private Const RULE_NAME_LISTPN  As String = "list_punctuation"

' -- Marker pattern types (Rule 10) ----------------------------
Private Const MARKER_LETTER As String = "letter"   ' (a), (b), (c)
Private Const MARKER_ROMAN  As String = "roman"    ' (i), (ii), (iii)
Private Const MARKER_NUMBER As String = "number"   ' (1), (2), (3)

' Max paragraphs between inline lists to consider them related
Private Const MAX_LIST_GAP As Long = 30

' ==============================================================
'  RULE 10 - PRIVATE HELPERS
' ==============================================================

' -- Helper: check if a parenthesized marker is a clause reference --
' Returns True if the opening paren is immediately preceded by a
' digit or letter (no space), e.g. "3(4)" or "Rule 1(a)" where
' the "1(" is adjacent. These are structural references, not lists.
Private Function IsClauseRef(ByRef paraText As String, _
                              ByVal openParen As Long) As Boolean
    Dim prevCh As String
    Dim refWords As Variant
    Dim wordEnd As Long
    Dim wStart As Long
    Dim wCh As String
    Dim prevWord As String
    Dim ri As Long
    Dim conjEnd As Long
    Dim conjWord As String
    Dim cStart As Long
    Dim cc As String
    Dim scanBack As Long

    IsClauseRef = False
    If openParen <= 1 Then Exit Function

    prevCh = Mid$(paraText, openParen - 1, 1)

    ' If preceded by a digit, letter, or closing paren -- clause ref
    If (prevCh >= "0" And prevCh <= "9") Or _
       (prevCh >= "A" And prevCh <= "Z") Or _
       (prevCh >= "a" And prevCh <= "z") Or _
       prevCh = ")" Then
        IsClauseRef = True
        Exit Function
    End If

    ' If not preceded by a space, nothing more to check
    If prevCh <> " " Then Exit Function

    ' -- Check for structural reference word before the space --
    refWords = Array("paragraph", "paragraphs", "para", "paras", _
                     "section", "sections", "sect", "sects", _
                     "clause", "clauses", "cl", _
                     "article", "articles", "art", "arts", _
                     "rule", "rules", "r", _
                     "regulation", "regulations", "reg", "regs", _
                     "schedule", "schedules", "sch", _
                     "sub-paragraph", "sub-paragraphs", _
                     "sub-section", "sub-sections", _
                     "sub-clause", "sub-clauses", _
                     "part", "parts", "pt", _
                     "item", "items", "annex")

    wordEnd = openParen - 2
    If wordEnd >= 1 Then
        wStart = wordEnd
        Do While wStart >= 1
            wCh = Mid$(paraText, wStart, 1)
            If (wCh >= "A" And wCh <= "Z") Or _
               (wCh >= "a" And wCh <= "z") Or wCh = "-" Then
                wStart = wStart - 1
            Else
                Exit Do
            End If
        Loop
        wStart = wStart + 1
        If wStart <= wordEnd Then
            prevWord = LCase(Mid$(paraText, wStart, wordEnd - wStart + 1))
            For ri = LBound(refWords) To UBound(refWords)
                If prevWord = CStr(refWords(ri)) Then
                    IsClauseRef = True
                    Exit Function
                End If
            Next ri
        End If
    End If

    ' -- Check for conjunction-linked clause ref --
    ' e.g. "paragraph (1) or (2)" -- the "(2)" preceded by "or "
    conjEnd = openParen - 2
    If conjEnd >= 1 Then
        cStart = conjEnd
        Do While cStart >= 1
            cc = Mid$(paraText, cStart, 1)
            If (cc >= "A" And cc <= "Z") Or (cc >= "a" And cc <= "z") Then
                cStart = cStart - 1
            Else
                Exit Do
            End If
        Loop
        cStart = cStart + 1
        If cStart <= conjEnd Then
            conjWord = LCase(Mid$(paraText, cStart, conjEnd - cStart + 1))
            If conjWord = "and" Or conjWord = "or" Or conjWord = "to" Then
                scanBack = cStart - 1
                Do While scanBack >= 1 And Mid$(paraText, scanBack, 1) = " "
                    scanBack = scanBack - 1
                Loop
                If scanBack >= 1 And Mid$(paraText, scanBack, 1) = ")" Then
                    IsClauseRef = True
                    Exit Function
                End If
            End If
        End If
    End If
End Function

' -- Helper: detect marker type from content between parens ----
Private Function GetMarkerType(ByVal content As String) As String
    If Len(content) = 0 Then
        GetMarkerType = ""
        Exit Function
    End If

    ' Single lowercase letter: (a)-(z)
    If Len(content) = 1 And content Like "[a-z]" Then
        GetMarkerType = MARKER_LETTER
        Exit Function
    End If

    ' Numeric: (1), (2), (12)
    If IsNumeric(content) Then
        GetMarkerType = MARKER_NUMBER
        Exit Function
    End If

    ' Roman numeral: all chars are i, v, x, l, c, d, m
    Dim allRoman As Boolean
    Dim ci As Long
    allRoman = True
    For ci = 1 To Len(content)
        If Not (Mid$(content, ci, 1) Like "[ivxlcdm]") Then
            allRoman = False
            Exit For
        End If
    Next ci
    If allRoman Then
        GetMarkerType = MARKER_ROMAN
        Exit Function
    End If

    GetMarkerType = ""
End Function

' -- Helper: find all inline list markers in a paragraph -------
' Returns Collection of Array(markerPos, markerText, markerContent, markerType)
Private Function FindMarkersInPara(ByVal paraText As String) As Collection
    Dim markers As New Collection
    Dim pos As Long
    Dim openParen As Long
    Dim closeParen As Long
    Dim content As String
    Dim info() As Variant
    Dim mType As String

    pos = 1
    Do While pos <= Len(paraText)
        openParen = InStr(pos, paraText, "(")
        If openParen = 0 Then Exit Do

        closeParen = InStr(openParen + 1, paraText, ")")
        If closeParen = 0 Then Exit Do
        If closeParen - openParen > 6 Then
            ' Too long to be a list marker
            pos = openParen + 1
            GoTo ContinueSearch
        End If

        content = Mid$(paraText, openParen + 1, closeParen - openParen - 1)
        mType = GetMarkerType(content)

        If Len(mType) > 0 And Not IsClauseRef(paraText, openParen) Then
            ReDim info(0 To 3)
            info(0) = openParen         ' position in paragraph text
            info(1) = Mid$(paraText, openParen, closeParen - openParen + 1) ' full marker text
            info(2) = content           ' content between parens
            info(3) = mType             ' marker type
            markers.Add info
        End If

        pos = closeParen + 1
ContinueSearch:
    Loop

    Set FindMarkersInPara = markers
End Function

' -- Helper: detect separator before a marker ------------------
' Looks at text between previous marker's end and current marker's start
Private Function DetectSeparator(ByVal textBetween As String) As String
    Dim trimmed As String
    trimmed = Trim$(textBetween)

    ' Check for semicolon
    If InStr(1, trimmed, ";") > 0 Then
        DetectSeparator = "semicolon"
        Exit Function
    End If

    ' Check for comma
    If InStr(1, trimmed, ",") > 0 Then
        DetectSeparator = "comma"
        Exit Function
    End If

    DetectSeparator = "none"
End Function

' -- Helper: check if conjunction precedes final marker --------
Private Function DetectConjunction(ByVal textBefore As String) As String
    Dim trimmed As String
    trimmed = LCase(Trim$(textBefore))

    ' Remove trailing semicolons/commas for checking
    Do While Len(trimmed) > 0 And (Right$(trimmed, 1) = ";" Or Right$(trimmed, 1) = ",")
        trimmed = Trim$(Left$(trimmed, Len(trimmed) - 1))
    Loop

    ' Check if ends with "and" or "or"
    If Len(trimmed) >= 3 Then
        If Right$(trimmed, 4) = " and" Or trimmed = "and" Then
            DetectConjunction = "and"
            Exit Function
        End If
        If Right$(trimmed, 3) = " or" Or trimmed = "or" Then
            DetectConjunction = "or"
            Exit Function
        End If
    End If

    DetectConjunction = "none"
End Function

' -- Helper: analyse one inline list paragraph and return its style key --
'  Returns "" if the paragraph is not a valid inline list.
Private Function AnalyseInlineList(ByVal paraText As String, _
        markers As Collection) As String

    ' Need at least 2 markers to form an inline list
    If markers.Count < 2 Then
        AnalyseInlineList = ""
        Exit Function
    End If

    ' Verify markers are of the same type and sequential
    Dim firstType As String
    Dim mk As Variant
    mk = markers(1)
    firstType = CStr(mk(3))

    Dim sameType As Boolean
    sameType = True
    Dim mi As Long
    For mi = 2 To markers.Count
        mk = markers(mi)
        If CStr(mk(3)) <> firstType Then
            sameType = False
            Exit For
        End If
    Next mi
    If Not sameType Then
        AnalyseInlineList = ""
        Exit Function
    End If

    ' -- Analyse separator style --------------------------
    Dim separators As New Collection
    For mi = 2 To markers.Count
        Dim prevMk As Variant
        prevMk = markers(mi - 1)
        Dim currMk As Variant
        currMk = markers(mi)

        Dim prevEnd As Long
        prevEnd = CLng(prevMk(0)) + Len(CStr(prevMk(1)))
        Dim currStart As Long
        currStart = CLng(currMk(0))

        If currStart > prevEnd Then
            Dim between As String
            between = Mid$(paraText, prevEnd, currStart - prevEnd)
            separators.Add DetectSeparator(between)
        Else
            separators.Add "none"
        End If
    Next mi

    ' Determine dominant separator for this list
    Dim sepSemi As Long, sepComma As Long, sepNone As Long
    sepSemi = 0: sepComma = 0: sepNone = 0
    Dim s As Variant
    For Each s In separators
        Select Case CStr(s)
            Case "semicolon": sepSemi = sepSemi + 1
            Case "comma": sepComma = sepComma + 1
            Case "none": sepNone = sepNone + 1
        End Select
    Next s

    Dim listSep As String
    If sepSemi >= sepComma And sepSemi >= sepNone Then
        listSep = "semicolon"
    ElseIf sepComma >= sepSemi And sepComma >= sepNone Then
        listSep = "comma"
    Else
        listSep = "none"
    End If

    ' -- Check conjunction before final marker ----------------
    Dim lastMk As Variant
    lastMk = markers(markers.Count)
    Dim lastMkStart As Long
    lastMkStart = CLng(lastMk(0))

    Dim secondLastMk As Variant
    secondLastMk = markers(markers.Count - 1)
    Dim slEnd As Long
    slEnd = CLng(secondLastMk(0)) + Len(CStr(secondLastMk(1)))

    Dim conjText As String
    If lastMkStart > slEnd Then
        conjText = Mid$(paraText, slEnd, lastMkStart - slEnd)
    Else
        conjText = ""
    End If
    Dim conjunction As String
    conjunction = DetectConjunction(conjText)

    ' -- Check ending punctuation -----------------------------
    Dim lastMkEnd As Long
    lastMkEnd = CLng(lastMk(0)) + Len(CStr(lastMk(1)))
    Dim afterLast As String
    If lastMkEnd <= Len(paraText) Then
        afterLast = Mid$(paraText, lastMkEnd)
    Else
        afterLast = ""
    End If
    Dim ending As String
    Dim cleanAfter As String
    cleanAfter = Trim$(Replace(afterLast, vbCr, ""))
    cleanAfter = Trim$(Replace(cleanAfter, vbLf, ""))
    If Len(cleanAfter) > 0 Then
        Dim lastChar As String
        lastChar = Right$(cleanAfter, 1)
        If lastChar = "." Then
            ending = "fullstop"
        ElseIf lastChar = ";" Then
            ending = "semicolon"
        Else
            ending = "none"
        End If
    Else
        ending = "none"
    End If

    ' -- Build style key --------------------------------------
    AnalyseInlineList = listSep & "|" & conjunction & "|" & ending
End Function

' ==============================================================
'  RULE 10 - PUBLIC FUNCTION: Check_InlineListFormat
'
'  LOCAL-CONTEXT APPROACH:
'  1. Collect all inline-list paragraphs with their style keys.
'  2. Group consecutive inline lists that are within MAX_LIST_GAP
'     paragraphs of each other into a "cluster".
'  3. Within each cluster, determine dominant style and flag
'     deviations.
' ==============================================================
Public Function Check_InlineListFormat(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' -- Collect all inline list paragraphs --------------------
    ' Each entry: Array(styleKey, paraIdx, rangeStart, rangeEnd, previewText)
    Dim listCap As Long
    listCap = 64
    Dim listCount As Long
    listCount = 0
    Dim lStyles() As String
    Dim lParaIdx() As Long
    Dim lStarts() As Long
    Dim lEnds() As Long
    Dim lPreviews() As String
    ReDim lStyles(0 To listCap - 1)
    ReDim lParaIdx(0 To listCap - 1)
    ReDim lStarts(0 To listCap - 1)
    ReDim lEnds(0 To listCap - 1)
    ReDim lPreviews(0 To listCap - 1)

    Dim para As Paragraph
    Dim paraIdx As Long

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        ' Page range filter
        If Not EngineIsInPageRange(para.Range) Then GoTo NextPara

        Dim paraText As String
        paraText = para.Range.Text
        If Err.Number <> 0 Then paraText = "": Err.Clear

        ' Find all markers in this paragraph
        Dim markers As Collection
        Set markers = FindMarkersInPara(paraText)

        ' Analyse as inline list
        Dim styleKey As String
        styleKey = AnalyseInlineList(paraText, markers)
        If Len(styleKey) = 0 Then GoTo NextPara

        ' Grow arrays if needed
        If listCount >= listCap Then
            listCap = listCap * 2
            ReDim Preserve lStyles(0 To listCap - 1)
            ReDim Preserve lParaIdx(0 To listCap - 1)
            ReDim Preserve lStarts(0 To listCap - 1)
            ReDim Preserve lEnds(0 To listCap - 1)
            ReDim Preserve lPreviews(0 To listCap - 1)
        End If

        lStyles(listCount) = styleKey
        lParaIdx(listCount) = paraIdx
        lStarts(listCount) = para.Range.Start
        If Err.Number <> 0 Then lStarts(listCount) = 0: Err.Clear
        lEnds(listCount) = para.Range.End
        If Err.Number <> 0 Then lEnds(listCount) = 0: Err.Clear
        lPreviews(listCount) = Trim$(Replace(Left$(paraText, 80), vbCr, ""))
        listCount = listCount + 1

NextPara:
    Next para

    If listCount < 2 Then
        On Error GoTo 0
        Set Check_InlineListFormat = issues
        Exit Function
    End If

    ' -- Group into local clusters ----------------------------
    ' A new cluster starts when the paragraph gap exceeds MAX_LIST_GAP
    Dim csCap As Long
    csCap = 16
    Dim csCount As Long
    csCount = 0
    Dim clusterStarts() As Long
    Dim clusterEnds() As Long
    ReDim clusterStarts(0 To csCap - 1)
    ReDim clusterEnds(0 To csCap - 1)

    Dim curClusterStart As Long
    curClusterStart = 0

    Dim li As Long
    For li = 1 To listCount - 1
        Dim gap As Long
        gap = lParaIdx(li) - lParaIdx(li - 1)
        If gap > MAX_LIST_GAP Then
            ' Close current cluster
            If csCount >= csCap Then
                csCap = csCap * 2
                ReDim Preserve clusterStarts(0 To csCap - 1)
                ReDim Preserve clusterEnds(0 To csCap - 1)
            End If
            clusterStarts(csCount) = curClusterStart
            clusterEnds(csCount) = li - 1
            csCount = csCount + 1
            curClusterStart = li
        End If
    Next li

    ' Close last cluster
    If csCount >= csCap Then
        csCap = csCap * 2
        ReDim Preserve clusterStarts(0 To csCap - 1)
        ReDim Preserve clusterEnds(0 To csCap - 1)
    End If
    clusterStarts(csCount) = curClusterStart
    clusterEnds(csCount) = listCount - 1
    csCount = csCount + 1

    ' -- Within each cluster, find dominant and flag -----------
    Dim ci As Long
    For ci = 0 To csCount - 1
        Dim cStart As Long
        cStart = clusterStarts(ci)
        Dim cEnd As Long
        cEnd = clusterEnds(ci)

        ' Need at least 2 lists in cluster to compare
        If cEnd - cStart < 1 Then GoTo NextCluster

        ' Count styles in this cluster
        Dim styleCounts As Object
        Set styleCounts = CreateObject("Scripting.Dictionary")
        Dim cj As Long
        For cj = cStart To cEnd
            If styleCounts.Exists(lStyles(cj)) Then
                styleCounts(lStyles(cj)) = styleCounts(lStyles(cj)) + 1
            Else
                styleCounts.Add lStyles(cj), 1
            End If
        Next cj

        ' Only flag if more than one style in this cluster
        If styleCounts.Count < 2 Then GoTo NextCluster

        ' Find dominant style
        Dim domStyle As String
        Dim maxCnt As Long
        domStyle = ""
        maxCnt = 0
        Dim sk As Variant
        For Each sk In styleCounts.keys
            If styleCounts(sk) > maxCnt Then
                maxCnt = styleCounts(sk)
                domStyle = CStr(sk)
            End If
        Next sk

        ' Flag deviations
        For cj = cStart To cEnd
            If lStyles(cj) <> domStyle Then
                Dim finding As Object
                Dim rng As Range
                Set rng = doc.Range(lStarts(cj), lEnds(cj))
                If Err.Number <> 0 Then Err.Clear: GoTo NextClusterItem
                Dim loc As String
                loc = EngineGetLocationString(rng, doc)
                If Err.Number <> 0 Then loc = "unknown location": Err.Clear

                ' Parse dominant style for suggestion
                Dim domParts() As String
                domParts = Split(domStyle, "|")
                Dim suggStr As String
                suggStr = "Use consistent list formatting: "
                If UBound(domParts) >= 0 Then suggStr = suggStr & domParts(0) & " separators"
                If UBound(domParts) >= 1 Then suggStr = suggStr & ", '" & domParts(1) & "' conjunction"
                If UBound(domParts) >= 2 Then suggStr = suggStr & ", " & domParts(2) & " ending"

                Set finding = CreateIssueDict(RULE_NAME_INLINE, loc, "Inline list format inconsistency near: '" & lPreviews(cj) & "...'", suggStr, lStarts(cj), lEnds(cj), "possible_error")
                issues.Add finding
            End If
NextClusterItem:
        Next cj
NextCluster:
    Next ci

    On Error GoTo 0
    Set Check_InlineListFormat = issues
End Function

' ==============================================================
'  RULE 15 - PRIVATE HELPERS
' ==============================================================

' -- Strip trailing carriage return / line feed ----------------
Private Function StripTrailingCr(ByVal text As String) As String
    Dim result As String
    result = text

    Do While Len(result) > 0
        Dim lastCh As String
        lastCh = Right(result, 1)
        If lastCh = vbCr Or lastCh = vbLf Or lastCh = Chr(13) Or lastCh = Chr(10) Then
            result = Left(result, Len(result) - 1)
        Else
            Exit Do
        End If
    Loop

    StripTrailingCr = result
End Function

' -- Get last N characters of a string -------------------------
Private Function GetLastNChars(ByVal text As String, ByVal n As Long) As String
    If Len(text) <= n Then
        GetLastNChars = text
    Else
        GetLastNChars = Right(text, n)
    End If
End Function

' -- Classify the ending punctuation of a list item ------------
Private Function ClassifyEnding(ByVal text As String) As String
    Dim trimmed As String
    Dim endChar As String

    trimmed = StripTrailingCr(text)
    trimmed = Trim(trimmed)

    If Len(trimmed) = 0 Then
        ClassifyEnding = "none"
        Exit Function
    End If

    endChar = Right(trimmed, 1)

    Select Case endChar
        Case ";"
            ClassifyEnding = "semicolon"
        Case "."
            ClassifyEnding = "full_stop"
        Case ","
            ClassifyEnding = "comma"
        Case ":"
            ClassifyEnding = "colon"
        Case Else
            ClassifyEnding = "none"
    End Select
End Function

' -- Process a single list group for punctuation issues --------
Private Sub ProcessListGroup(doc As Document, _
                              ByRef issues As Collection, _
                              ByRef paraStarts() As Long, _
                              ByRef paraEnds() As Long, _
                              ByRef paraTexts() As String, _
                              ByVal groupStart As Long, _
                              ByVal groupEnd As Long)
    Dim itemCount As Long
    Dim i As Long
    Dim endings() As String
    Dim endingCounts As Object ' Dictionary
    Dim dominantEnding As String
    Dim maxCount As Long

    itemCount = groupEnd - groupStart + 1
    If itemCount < 2 Then Exit Sub ' Single-item list, nothing to check

    ' -- Classify the ending of each list item ------------------
    ReDim endings(groupStart To groupEnd)

    For i = groupStart To groupEnd
        endings(i) = ClassifyEnding(paraTexts(i))
    Next i

    ' -- Count endings to find dominant -------------------------
    Set endingCounts = CreateObject("Scripting.Dictionary")

    For i = groupStart To groupEnd
        If endingCounts.Exists(endings(i)) Then
            endingCounts(endings(i)) = endingCounts(endings(i)) + 1
        Else
            endingCounts.Add endings(i), 1
        End If
    Next i

    dominantEnding = ""
    maxCount = 0
    Dim key As Variant
    For Each key In endingCounts.keys
        If endingCounts(key) > maxCount Then
            maxCount = endingCounts(key)
            dominantEnding = CStr(key)
        End If
    Next key

    ' -- Flag items that deviate from dominant ending ------------
    For i = groupStart To groupEnd
        If endings(i) <> dominantEnding Then
            ' Skip the last item if dominant is semicolon (special rule below)
            If dominantEnding = "semicolon" And i = groupEnd Then
                GoTo ContinueItem
            End If

            Dim rng As Range
            Dim locStr As String
            Dim finding As Object

            On Error Resume Next
            Set rng = doc.Range(paraStarts(i), paraEnds(i))
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo ContinueItem
            End If

            If Not EngineIsInPageRange(rng) Then
                On Error GoTo 0
                GoTo ContinueItem
            End If

            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_LISTPN, locStr, "List item ending '" & endings(i) & "' differs from " & "dominant ending '" & dominantEnding & "'", "Change ending punctuation to match list style (" & dominantEnding & ")", paraStarts(i), paraEnds(i), "possible_error")
            issues.Add finding
        End If

ContinueItem:
    Next i

    ' -- Special: if dominant is semicolon, last item should end with full stop -
    If dominantEnding = "semicolon" Then
        If endings(groupEnd) <> "full_stop" Then
            On Error Resume Next
            Set rng = doc.Range(paraStarts(groupEnd), paraEnds(groupEnd))
            If Err.Number = 0 Then
                If EngineIsInPageRange(rng) Then
                    locStr = EngineGetLocationString(rng, doc)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    End If

                    Set finding = CreateIssueDict(RULE_NAME_LISTPN, locStr, "Last list item should end with a full stop, not '" & endings(groupEnd) & "'", "End the final list item with a full stop", paraStarts(groupEnd), paraEnds(groupEnd), "possible_error")
                    issues.Add finding
                End If
            End If
            On Error GoTo 0
        End If

        ' -- Check penultimate item for "and" or "or" -----------
        If itemCount >= 2 Then
            Dim penIdx As Long
            penIdx = groupEnd - 1
            Dim penText As String
            penText = LCase(Trim(StripTrailingCr(paraTexts(penIdx))))

            Dim hasConjunction As Boolean
            hasConjunction = False

            ' Check if text ends with "and;" or "or;" or similar
            If Right(penText, 4) = "and;" Or Right(penText, 3) = "or;" Or _
               Right(penText, 4) = "and," Or Right(penText, 3) = "or," Or _
               Right(penText, 3) = "and" Or Right(penText, 2) = "or" Then
                hasConjunction = True
            End If

            ' Also check for "and" / "or" as last word before punctuation
            Dim lastWords As String
            lastWords = GetLastNChars(penText, 10)
            If InStr(1, lastWords, " and") > 0 Or InStr(1, lastWords, " or") > 0 Then
                hasConjunction = True
            End If

            If Not hasConjunction Then
                On Error Resume Next
                Set rng = doc.Range(paraStarts(penIdx), paraEnds(penIdx))
                If Err.Number = 0 Then
                    If EngineIsInPageRange(rng) Then
                        locStr = EngineGetLocationString(rng, doc)
                        If Err.Number <> 0 Then
                            locStr = "unknown location"
                            Err.Clear
                        End If

                        Set finding = CreateIssueDict(RULE_NAME_LISTPN, locStr, "Penultimate list item should include 'and' or 'or' " & "before terminal punctuation", "Add 'and' or 'or' before the semicolon", paraStarts(penIdx), paraEnds(penIdx), "possible_error")
                        issues.Add finding
                    End If
                End If
                On Error GoTo 0
            End If
        End If
    End If
End Sub

' ==============================================================
'  RULE 15 - PUBLIC FUNCTION: Check_ListPunctuation
' ==============================================================
Public Function Check_ListPunctuation(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim totalParas As Long

    ' -- Collect all paragraphs into arrays for easier processing -
    totalParas = doc.Paragraphs.Count
    If totalParas = 0 Then
        Set Check_ListPunctuation = issues
        Exit Function
    End If

    ' Arrays to hold paragraph info
    Dim paraStarts() As Long
    Dim paraEnds() As Long
    Dim paraTexts() As String
    Dim paraIsList() As Boolean
    Dim paraListID() As Long

    ReDim paraStarts(1 To totalParas)
    ReDim paraEnds(1 To totalParas)
    ReDim paraTexts(1 To totalParas)
    ReDim paraIsList(1 To totalParas)
    ReDim paraListID(1 To totalParas)

    paraIdx = 0
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        paraIdx = paraIdx + 1

        Dim paraRange As Range
        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            paraStarts(paraIdx) = 0
            paraEnds(paraIdx) = 0
            paraTexts(paraIdx) = ""
            paraIsList(paraIdx) = False
            paraListID(paraIdx) = 0
            GoTo NextParaCollect
        End If

        paraStarts(paraIdx) = paraRange.Start
        paraEnds(paraIdx) = paraRange.End
        paraTexts(paraIdx) = paraRange.Text

        ' Check if paragraph is a list item
        Dim listType As Long
        listType = 0
        listType = paraRange.ListFormat.ListType
        If Err.Number <> 0 Then
            Err.Clear
            listType = 0
        End If

        paraIsList(paraIdx) = (listType <> 0) ' 0 = wdListNoNumbering

        ' Get a list identifier for grouping (start pos of first para
        ' in the Word List object -- unique per list in the document)
        Dim listID As Long
        listID = 0
        If paraIsList(paraIdx) Then
            listID = paraRange.ListFormat.List.ListParagraphs(1).Range.Start
            If Err.Number <> 0 Then
                Err.Clear
                ' Fallback: use list level + approximate position
                listID = paraRange.ListFormat.ListLevelNumber + 1
                If Err.Number <> 0 Then
                    Err.Clear
                    listID = 0  ' unknown -- do not use for group-breaking
                End If
            End If
        End If
        paraListID(paraIdx) = listID

NextParaCollect:
    Next para
    On Error GoTo 0

    ' -- Group consecutive list paragraphs into lists -----------
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim inGroup As Boolean

    inGroup = False
    Dim p As Long

    For p = 1 To totalParas
        If paraIsList(p) Then
            If Not inGroup Then
                groupStart = p
                inGroup = True
            ElseIf paraListID(p) <> 0 And paraListID(groupStart) <> 0 _
                   And paraListID(p) <> paraListID(groupStart) Then
                ' Different list -- close current group, start new one
                ProcessListGroup doc, issues, paraStarts, paraEnds, paraTexts, _
                                 groupStart, groupEnd
                groupStart = p
            End If
            groupEnd = p
        Else
            If inGroup Then
                ' Process the list group
                ProcessListGroup doc, issues, paraStarts, paraEnds, paraTexts, _
                                 groupStart, groupEnd
                inGroup = False
            End If
        End If
    Next p

    ' Process final group if document ends with a list
    If inGroup Then
        ProcessListGroup doc, issues, paraStarts, paraEnds, paraTexts, _
                         groupStart, groupEnd
    End If

    Set Check_ListPunctuation = issues
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_NumberFormats.bas

```vb
Attribute VB_Name = "Rules_NumberFormats"
' ============================================================
' Rules_NumberFormats.bas
' Combined module for number/date/currency format rules:
'   - Rule09: Date and time format consistency
'   - Rule19: Currency and number format consistency
'
' RETIRED (not engine-wired):
'   - Rule18 page-range helpers: kept for backwards compatibility
'     but not dispatched by RunAllPleadingsRules. The engine
'     manages page ranges directly via SetPageRangeFromString.
'
' Public functions:
'   Check_DateTimeFormat        (Rule09)
'   Check_CurrencyNumberFormat  (Rule19)
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

' -- Rule name constants ---------------------------------------
Private Const RULE_NAME_DATE_TIME As String = "date_time_format"
' RETIRED -- DEAD CODE: page_range is not engine-wired and this constant is unused.
' Kept only so the module compiles if an external caller references it.
Private Const RETIRED_RULE_NAME_PAGE_RANGE As String = "page_range"
Private Const RULE_NAME_CURRENCY As String = "currency_number_format"

' -- Currency format category constants (Rule19) ---------------
Private Const FMT_WORDS As String = "words"
Private Const FMT_ABBREVIATED As String = "abbreviated"
Private Const FMT_FULL_NUMERIC As String = "full_numeric"
Private Const FMT_ISO_PREFIX As String = "iso_prefix"

' -- Module-level page range state (Rule18) --------------------
Private mStartPage As Long   ' 0 = no restriction
Private mEndPage   As Long   ' 0 = no restriction

' ============================================================
'  PRIVATE HELPERS  -  Rule09 (Date/Time)
' ============================================================

' -- Helper: validate a month name -----------------------------
Private Function IsValidMonth(ByVal monthName As String) As Boolean
    Dim months As Variant
    months = Array("January", "February", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")
    Dim m As Variant
    For Each m In months
        If StrComp(monthName, CStr(m), vbTextCompare) = 0 Then
            IsValidMonth = True
            Exit Function
        End If
    Next m
    IsValidMonth = False
End Function

' -- Helper: search and collect date/time occurrences ----------
Private Sub FindWithWildcard(doc As Document, ByVal pattern As String, _
                              results As Collection, ByVal formatType As String)
    Dim rng As Range
    Dim info() As Variant
    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = pattern
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .MatchCase = False
    End With

    Dim lastPos As Long
    lastPos = -1
    Do While rng.Find.Execute
        If rng.Start <= lastPos Then Exit Do  ' stall guard
        lastPos = rng.Start
        If EngineIsInPageRange(rng) Then
            ReDim info(0 To 3)
            info(0) = formatType
            info(1) = rng.Text
            info(2) = rng.Start
            info(3) = rng.End
            results.Add info
        End If
        rng.Collapse wdCollapseEnd
    Loop
End Sub

' -- Helper: check if a time match looks like a clause reference,
'  ratio, date component, or other non-time pattern.
'  Examines characters before and after the HH:MM match.
' ----------------------------------------------------------------
Private Function LooksLikeNonTimeContext(doc As Document, _
        ByVal matchStart As Long, ByVal matchEnd As Long) As Boolean
    LooksLikeNonTimeContext = False
    On Error Resume Next

    ' Check character before the match
    If matchStart > 0 Then
        Dim bRng As Range
        Set bRng = doc.Range(matchStart - 1, matchStart)
        If Err.Number = 0 Then
            Dim bc As String
            bc = bRng.Text
            If Err.Number = 0 Then
                ' Preceded by letter -> probably part of a word or reference
                If (bc >= "A" And bc <= "Z") Or (bc >= "a" And bc <= "z") Then
                    LooksLikeNonTimeContext = True
                    Err.Clear: On Error GoTo 0: Exit Function
                End If
                ' Preceded by another digit -> could be ratio like 1:12:45
                If bc >= "0" And bc <= "9" Then
                    ' Check two chars back for another colon (chained ratio)
                    If matchStart > 1 Then
                        Dim b2Rng As Range
                        Set b2Rng = doc.Range(matchStart - 2, matchStart - 1)
                        If Err.Number = 0 Then
                            Dim b2c As String
                            b2c = b2Rng.Text
                            If b2c = ":" Or b2c = "." Then
                                LooksLikeNonTimeContext = True
                                Err.Clear: On Error GoTo 0: Exit Function
                            End If
                        Else
                            Err.Clear
                        End If
                    End If
                End If
            Else
                Err.Clear
            End If
        Else
            Err.Clear
        End If
    End If

    ' Check character after the match
    If matchEnd < doc.Content.End Then
        Dim aRng As Range
        Set aRng = doc.Range(matchEnd, matchEnd + 1)
        If Err.Number = 0 Then
            Dim ac As String
            ac = aRng.Text
            If Err.Number = 0 Then
                ' Followed by a colon or dot+digit -> ratio or version number
                If ac = ":" Then
                    LooksLikeNonTimeContext = True
                    Err.Clear: On Error GoTo 0: Exit Function
                End If
                If ac = "." Then
                    If matchEnd + 1 < doc.Content.End Then
                        Dim a2Rng As Range
                        Set a2Rng = doc.Range(matchEnd + 1, matchEnd + 2)
                        If Err.Number = 0 Then
                            Dim a2c As String
                            a2c = a2Rng.Text
                            If a2c >= "0" And a2c <= "9" Then
                                LooksLikeNonTimeContext = True
                                Err.Clear: On Error GoTo 0: Exit Function
                            End If
                        Else
                            Err.Clear
                        End If
                    End If
                End If
                ' Followed by a letter -> part of a word
                If (ac >= "A" And ac <= "Z") Or (ac >= "a" And ac <= "z") Then
                    LooksLikeNonTimeContext = True
                    Err.Clear: On Error GoTo 0: Exit Function
                End If
            Else
                Err.Clear
            End If
        Else
            Err.Clear
        End If
    End If

    On Error GoTo 0
End Function

' ============================================================
'  PRIVATE HELPERS  -  Rule19 (Currency/Number)
' ============================================================

' -- Check format consistency for a single symbol --------------
'  Searches for words, abbreviated, and full_numeric formats,
'  determines the dominant format, and flags minorities.
Private Sub CheckSymbolConsistency(doc As Document, _
                                    sym As String, _
                                    symLabel As String, _
                                    ByRef issues As Collection)
    Dim wordsCount As Long
    Dim abbrCount As Long
    Dim numericCount As Long
    Dim wordsRanges As Collection
    Dim abbrRanges As Collection
    Dim numericRanges As Collection

    Set wordsRanges = New Collection
    Set abbrRanges = New Collection
    Set numericRanges = New Collection

    ' -- Search for "words" format: symbol + digits + space + word --
    ' Pattern: e.g. ?[0-9.]@ [a-z]@  (wildcard)
    Dim rng As Range
    Dim wordPattern As String
    wordPattern = sym & "[0-9.]@" & " [a-z]@"

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = wordPattern
        .MatchWildcards = True
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    Do
        On Error Resume Next
        Dim found As Boolean
        found = rng.Find.Execute
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Do
        End If
        On Error GoTo 0

        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        ' Validate that the trailing word is a magnitude word
        Dim matchText As String
        matchText = LCase(rng.Text)
        If IsMagnitudeWord(matchText) Then
            If EngineIsInPageRange(rng) Then
                wordsCount = wordsCount + 1
                wordsRanges.Add doc.Range(rng.Start, rng.End)
            End If
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    ' -- Search for "abbreviated" format: symbol + digits + m/bn/k --
    Dim abbrPattern As String
    abbrPattern = sym & "[0-9.]@[mbk]"

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = abbrPattern
        .MatchWildcards = True
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    lastPos = -1
    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        If EngineIsInPageRange(rng) Then
            abbrCount = abbrCount + 1
            abbrRanges.Add doc.Range(rng.Start, rng.End)
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    ' -- Search for "full_numeric" format: symbol + digits with commas --
    Dim numPattern As String
    numPattern = sym & "[0-9,.]@"

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = numPattern
        .MatchWildcards = True
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    lastPos = -1
    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        ' Only count as full_numeric if it contains a comma and is long enough
        Dim numText As String
        numText = rng.Text
        If InStr(numText, ",") > 0 And Len(numText) >= 5 Then
            If EngineIsInPageRange(rng) Then
                numericCount = numericCount + 1
                numericRanges.Add doc.Range(rng.Start, rng.End)
            End If
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    ' -- Determine dominant format and flag minorities ----------
    Dim totalFormats As Long
    totalFormats = 0
    If wordsCount > 0 Then totalFormats = totalFormats + 1
    If abbrCount > 0 Then totalFormats = totalFormats + 1
    If numericCount > 0 Then totalFormats = totalFormats + 1

    ' Only flag if more than one format is in use
    If totalFormats < 2 Then Exit Sub

    ' Find the dominant format
    Dim domFormat As String
    Dim domCount As Long
    domFormat = FMT_WORDS: domCount = wordsCount
    If abbrCount > domCount Then domFormat = FMT_ABBREVIATED: domCount = abbrCount
    If numericCount > domCount Then domFormat = FMT_FULL_NUMERIC: domCount = numericCount

    ' Flag minority: words
    If wordsCount > 0 And domFormat <> FMT_WORDS Then
        FlagMinorityRanges doc, wordsRanges, symLabel, FMT_WORDS, domFormat, issues
    End If

    ' Flag minority: abbreviated
    If abbrCount > 0 And domFormat <> FMT_ABBREVIATED Then
        FlagMinorityRanges doc, abbrRanges, symLabel, FMT_ABBREVIATED, domFormat, issues
    End If

    ' Flag minority: full_numeric
    If numericCount > 0 And domFormat <> FMT_FULL_NUMERIC Then
        FlagMinorityRanges doc, numericRanges, symLabel, FMT_FULL_NUMERIC, domFormat, issues
    End If
End Sub

' -- Check ISO code prefixed amounts ---------------------------
'  Searches for patterns like "GBP 1,500" or "USD 25.00"
Private Sub CheckISOCodeFormat(doc As Document, _
                                isoCode As String, _
                                ByRef issues As Collection)
    Dim rng As Range
    Dim isoPattern As String
    Dim finding As Object
    Dim locStr As String

    ' Search for ISO code followed by space and number
    isoPattern = isoCode & " [0-9]@"

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = isoPattern
        .MatchWildcards = True
        .MatchCase = True
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim isoCount As Long
    isoCount = 0
    Dim isoLastPos As Long
    isoLastPos = -1

    Do
        On Error Resume Next
        Dim isoFound As Boolean
        isoFound = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not isoFound Then Exit Do
        If rng.Start <= isoLastPos Then Exit Do   ' stall guard
        isoLastPos = rng.Start

        If EngineIsInPageRange(rng) Then
            isoCount = isoCount + 1

            ' Flag ISO prefix usage as informational (possible_error)
            ' since mixing ISO codes with symbol notation is inconsistent
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_CURRENCY, locStr, "ISO code format used: '" & rng.Text & "'", "Consider using symbol notation for consistency", rng.Start, rng.End, "possible_error")
            issues.Add finding
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' -- Flag all ranges in a minority format collection -----------
Private Sub FlagMinorityRanges(doc As Document, _
                                ranges As Collection, _
                                symLabel As String, _
                                minorityFmt As String, _
                                dominantFmt As String, _
                                ByRef issues As Collection)
    Dim i As Long
    Dim rng As Range
    Dim finding As Object
    Dim locStr As String

    For i = 1 To ranges.Count
        Set rng = ranges(i)

        On Error Resume Next
        locStr = EngineGetLocationString(rng, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
        On Error GoTo 0

        Set finding = CreateIssueDict(RULE_NAME_CURRENCY, locStr, symLabel & " amount uses '" & minorityFmt & "' format: '" & rng.Text & "'", "Use '" & dominantFmt & "' format for consistency (dominant style)", rng.Start, rng.End, "error")
        issues.Add finding
    Next i
End Sub

' -- Check if matched text contains a magnitude word -----------
Private Function IsMagnitudeWord(ByVal txt As String) As Boolean
    Dim lTxt As String
    lTxt = LCase(txt)

    IsMagnitudeWord = (InStr(lTxt, "million") > 0) Or _
                      (InStr(lTxt, "billion") > 0) Or _
                      (InStr(lTxt, "thousand") > 0) Or _
                      (InStr(lTxt, "hundred") > 0) Or _
                      (InStr(lTxt, "trillion") > 0)
End Function

' ============================================================
'  PUBLIC FUNCTIONS
' ============================================================

' ================================================================
'  Rule09: Check_DateTimeFormat
'  Detects date and time format inconsistencies across the
'  document. Identifies UK, US, and numeric date formats,
'  determines the dominant style, and flags deviations.
'  Also checks for mixed 12-hour / 24-hour time formats.
'
'  24-hour detection recognises 00:00 through 23:59 with
'  context filtering to exclude clause references and ratios.
' ================================================================
Public Function Check_DateTimeFormat(doc As Document) As Collection
    Dim issues As New Collection

    ' ==========================================================
    '  PASS 1: Find all date occurrences
    ' ==========================================================
    Dim dateFinds As New Collection
    Dim dateCounts As Object
    Set dateCounts = CreateObject("Scripting.Dictionary")
    dateCounts.Add "UK", 0
    dateCounts.Add "US", 0
    dateCounts.Add "numeric", 0

    ' -- UK format: "1 January 2024" or "12 March 2025" ------
    ' VBA wildcard: one or two digits, space, word, space, four digits
    Dim ukResults As New Collection
    FindWithWildcard doc, "[0-9]{1,2} [A-Z][a-z]{2,} [0-9]{4}", ukResults, "UK"

    ' Validate UK results (check month name)
    Dim ukItem As Variant
    Dim i As Long
    For i = 1 To ukResults.Count
        Dim ukInfo As Variant
        ukInfo = ukResults(i)
        Dim ukText As String
        ukText = CStr(ukInfo(1))

        ' Extract month name (between first and last space)
        Dim parts() As String
        parts = Split(ukText, " ")
        If UBound(parts) >= 2 Then
            If IsValidMonth(parts(1)) Then
                dateFinds.Add ukInfo
                dateCounts("UK") = dateCounts("UK") + 1
            End If
        End If
    Next i

    ' -- US format: "January 1, 2024" or "March 12, 2025" ----
    Dim usResults As New Collection
    FindWithWildcard doc, "[A-Z][a-z]{2,} [0-9]{1,2}, [0-9]{4}", usResults, "US"

    For i = 1 To usResults.Count
        Dim usInfo As Variant
        usInfo = usResults(i)
        Dim usText As String
        usText = CStr(usInfo(1))

        parts = Split(usText, " ")
        If UBound(parts) >= 0 Then
            If IsValidMonth(parts(0)) Then
                dateFinds.Add usInfo
                dateCounts("US") = dateCounts("US") + 1
            End If
        End If
    Next i

    ' -- Numeric format: "01/02/2024" or "1/2/24" -------------
    Dim numResults As New Collection
    FindWithWildcard doc, "[0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4}", numResults, "numeric"

    For i = 1 To numResults.Count
        dateFinds.Add numResults(i)
        dateCounts("numeric") = dateCounts("numeric") + 1
    Next i

    ' -- Determine dominant date format ------------------------
    Dim dominantDate As String
    Dim maxDateCount As Long
    Dim dk As Variant

    ' Check user preference first
    Dim datePref As String
    datePref = EngineGetDateFormatPref()

    If datePref = "UK" Or datePref = "US" Then
        ' User has set a preference -- use it as dominant
        dominantDate = datePref
        maxDateCount = dateCounts(datePref)
    Else
        ' AUTO mode: pick the most frequent format
        dominantDate = ""
        maxDateCount = 0
        For Each dk In dateCounts.keys
            If dateCounts(dk) > maxDateCount Then
                maxDateCount = dateCounts(dk)
                dominantDate = CStr(dk)
            End If
        Next dk
    End If

    ' -- Flag non-dominant date formats ------------------------
    If maxDateCount > 0 Then
        Dim totalDateFormats As Long
        totalDateFormats = 0
        For Each dk In dateCounts.keys
            If dateCounts(dk) > 0 Then totalDateFormats = totalDateFormats + 1
        Next dk

        ' Flag if there are mixed formats, or if a preference is set
        If totalDateFormats > 1 Or (datePref = "UK" Or datePref = "US") Then
            For i = 1 To dateFinds.Count
                Dim dInfo As Variant
                dInfo = dateFinds(i)
                Dim dType As String
                dType = CStr(dInfo(0))

                If dType <> dominantDate Then
                    Dim findingD As Object
                    Dim rngD As Range
                    On Error Resume Next
                    Set rngD = doc.Range(CLng(dInfo(2)), CLng(dInfo(3)))
                    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextDateFind
                    Dim locD As String
                    locD = EngineGetLocationString(rngD, doc)
                    If Err.Number <> 0 Then locD = "unknown location": Err.Clear
                    On Error GoTo 0

                    Dim suggestion As String
                    Select Case dominantDate
                        Case "UK"
                            suggestion = "Reformat to UK style (e.g., '1 January 2024')"
                        Case "US"
                            suggestion = "Reformat to US style (e.g., 'January 1, 2024')"
                        Case "numeric"
                            suggestion = "Reformat to numeric style (e.g., '01/01/2024')"
                    End Select

                    Set findingD = CreateIssueDict(RULE_NAME_DATE_TIME, locD, "Inconsistent date format: '" & CStr(dInfo(1)) & "' uses " & dType & " format but dominant is " & dominantDate, suggestion, CLng(dInfo(2)), CLng(dInfo(3)), "error")
                    issues.Add findingD
                End If
NextDateFind:
            Next i
        End If
    End If

    ' ==========================================================
    '  PASS 2: Find time format inconsistencies
    '
    '  12-hour: explicit AM/PM marker (e.g. 2:30 PM, 11:00 am)
    '  24-hour: HH:MM where HH is 00-23, no AM/PM follows,
    '           and context does not suggest clause ref or ratio.
    ' ==========================================================
    Dim timeFinds As New Collection
    Dim timeCounts As Object
    Set timeCounts = CreateObject("Scripting.Dictionary")
    timeCounts.Add "12hr", 0
    timeCounts.Add "24hr", 0

    ' -- 12-hour format: "2:30 PM", "11:00 am" ----------------
    Dim time12Results As New Collection
    FindWithWildcard doc, "[0-9]{1,2}:[0-9]{2} [AaPp][Mm]", time12Results, "12hr"

    For i = 1 To time12Results.Count
        timeFinds.Add time12Results(i)
        timeCounts("12hr") = timeCounts("12hr") + 1
    Next i

    ' Also catch dot-separated 12hr times: "2.30 pm"
    Dim time12DotResults As New Collection
    FindWithWildcard doc, "[0-9]{1,2}.[0-9]{2} [AaPp][Mm]", time12DotResults, "12hr"

    For i = 1 To time12DotResults.Count
        timeFinds.Add time12DotResults(i)
        timeCounts("12hr") = timeCounts("12hr") + 1
    Next i

    ' -- 24-hour format: HH:MM (00:00 through 23:59) ----------
    '  Search for two-digit colon two-digit patterns.
    '  Filter: must be valid 00-23 hour and 00-59 minute.
    '  Exclude matches followed by AM/PM (those are 12-hour).
    '  Exclude matches in non-time context (clause refs, ratios).
    Dim time24Results As New Collection
    FindWithWildcard doc, "[0-9]{2}:[0-9]{2}", time24Results, "24hr"

    For i = 1 To time24Results.Count
        Dim t24Info As Variant
        t24Info = time24Results(i)
        Dim t24Text As String
        t24Text = CStr(t24Info(1))

        ' Parse hour and minute
        Dim colonPos As Long
        colonPos = InStr(1, t24Text, ":")
        If colonPos > 0 Then
            Dim hourStr As String
            hourStr = Left$(t24Text, colonPos - 1)
            Dim minStr As String
            minStr = Mid$(t24Text, colonPos + 1)
            Dim hourVal As Long
            Dim minVal As Long
            hourVal = -1
            minVal = -1
            If IsNumeric(hourStr) Then hourVal = CLng(hourStr)
            If IsNumeric(minStr) Then minVal = CLng(minStr)

            ' Valid time: hour 0-23, minute 0-59
            If hourVal >= 0 And hourVal <= 23 And minVal >= 0 And minVal <= 59 Then
                Dim is24hrTime As Boolean
                is24hrTime = True

                ' Check whether AM/PM follows (with or without space)
                ' to avoid double-counting 12-hour times
                Dim peekEnd As Long
                peekEnd = CLng(t24Info(3)) + 4
                On Error Resume Next
                If peekEnd > doc.Content.End Then peekEnd = doc.Content.End
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                If peekEnd > CLng(t24Info(3)) Then
                    Dim peekRng As Range
                    On Error Resume Next
                    Set peekRng = doc.Range(CLng(t24Info(3)), peekEnd)
                    Dim peekTxt As String
                    peekTxt = ""
                    peekTxt = UCase$(peekRng.Text)
                    If Err.Number <> 0 Then peekTxt = "": Err.Clear
                    On Error GoTo 0
                    ' Followed by AM/PM (with or without space) = 12-hour
                    If Len(peekTxt) >= 2 Then
                        If Left$(peekTxt, 2) = "AM" Or Left$(peekTxt, 2) = "PM" Then
                            is24hrTime = False
                        ElseIf Len(peekTxt) >= 3 Then
                            If Mid$(peekTxt, 2, 2) = "AM" Or Mid$(peekTxt, 2, 2) = "PM" Then
                                is24hrTime = False
                            End If
                        End If
                    End If
                End If

                ' Context check: exclude clause refs, ratios, etc.
                If is24hrTime Then
                    If LooksLikeNonTimeContext(doc, CLng(t24Info(2)), CLng(t24Info(3))) Then
                        is24hrTime = False
                    End If
                End If

                ' Classify: hours 13-23 or 00 are definite 24-hour.
                ' Hours 01-12 without AM/PM are ambiguous and should not
                ' drive the dominant-style count (but are still collected
                ' so they can be flagged if a clear dominant emerges).
                If is24hrTime Then
                    If hourVal >= 13 Or hourVal = 0 Then
                        ' Definite 24-hour: counts toward dominance
                        timeFinds.Add t24Info
                        timeCounts("24hr") = timeCounts("24hr") + 1
                    Else
                        ' Ambiguous (01:00-12:59 without AM/PM):
                        ' Collect for possible flagging but mark as "ambiguous"
                        ' so it does NOT influence the dominant format.
                        Dim ambigInfo(0 To 3) As Variant
                        ambigInfo(0) = "ambiguous"
                        ambigInfo(1) = t24Info(1)
                        ambigInfo(2) = t24Info(2)
                        ambigInfo(3) = t24Info(3)
                        timeFinds.Add ambigInfo
                        ' Do NOT increment timeCounts("24hr")
                    End If
                End If
            End If
        End If
    Next i

    ' -- Determine dominant time format and flag deviations ----
    Dim dominantTime As String
    Dim maxTimeCount As Long
    dominantTime = ""
    maxTimeCount = 0
    For Each dk In timeCounts.keys
        If timeCounts(dk) > maxTimeCount Then
            maxTimeCount = timeCounts(dk)
            dominantTime = CStr(dk)
        End If
    Next dk

    If maxTimeCount > 0 Then
        Dim totalTimeFormats As Long
        totalTimeFormats = 0
        For Each dk In timeCounts.keys
            If timeCounts(dk) > 0 Then totalTimeFormats = totalTimeFormats + 1
        Next dk

        If totalTimeFormats > 1 Then
            For i = 1 To timeFinds.Count
                Dim tInfo As Variant
                tInfo = timeFinds(i)
                Dim tType As String
                tType = CStr(tInfo(0))

                ' Skip ambiguous times: they don't conflict with anything
                If tType = "ambiguous" Then GoTo NextTimeFind
                If tType <> dominantTime Then
                    Dim findingT As Object
                    Dim rngT As Range
                    On Error Resume Next
                    Set rngT = doc.Range(CLng(tInfo(2)), CLng(tInfo(3)))
                    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextTimeFind
                    Dim locT As String
                    locT = EngineGetLocationString(rngT, doc)
                    If Err.Number <> 0 Then locT = "unknown location": Err.Clear
                    On Error GoTo 0

                    Dim timeSugg As String
                    If dominantTime = "12hr" Then
                        timeSugg = "Use 12-hour format (e.g., '2:30 PM') for consistency"
                    Else
                        timeSugg = "Use 24-hour format (e.g., '14:30') for consistency"
                    End If

                    Set findingT = CreateIssueDict(RULE_NAME_DATE_TIME, locT, "Inconsistent time format: '" & CStr(tInfo(1)) & "' uses " & tType & " format but dominant is " & dominantTime, timeSugg, CLng(tInfo(2)), CLng(tInfo(3)), "error")
                    issues.Add findingT
                End If
NextTimeFind:
            Next i
        End If
    End If

    Set Check_DateTimeFormat = issues
End Function

' ================================================================
'  RETIRED Rule18: SetRange
'  NOT dispatched by the engine. The engine manages page ranges
'  directly via SetPageRangeFromString / SetPageRange.
'  Kept ONLY for backwards compatibility if called externally.
'  Will emit a debug warning when invoked.
' ================================================================
Public Sub SetRange(s As Long, e As Long)
    Debug.Print "WARNING: Rules_NumberFormats.SetRange is RETIRED (Rule18). " & _
                "Use PleadingsEngine.SetPageRange instead."
    mStartPage = s
    mEndPage = e
End Sub

' ================================================================
'  RETIRED Rule18: Check_PageRange
'  NOT dispatched by the engine. The engine manages page ranges
'  directly via SetPageRangeFromString / SetPageRange.
'  Kept ONLY for backwards compatibility if called externally.
'  Returns an empty collection; will emit a debug warning.
' ================================================================
Public Function Check_PageRange(doc As Document) As Collection
    Debug.Print "WARNING: Rules_NumberFormats.Check_PageRange is RETIRED (Rule18). " & _
                "Not dispatched by RunAllPleadingsRules."
    Dim issues As New Collection

    On Error Resume Next

    ' Push the stored page range into the engine
    EngineSetPageRange mStartPage, mEndPage

    On Error GoTo 0

    Set Check_PageRange = issues
End Function

' ================================================================
'  Rule19: Check_CurrencyNumberFormat
'  Detects inconsistent currency/number formatting across
'  the document. Checks symbol-prefixed amounts (GBP, USD, EUR)
'  and ISO-code-prefixed amounts, then flags minority format
'  usage.
' ================================================================
Public Function Check_CurrencyNumberFormat(doc As Document) As Collection
    Dim issues As New Collection
    Dim symbols As Variant
    Dim symLabels As Variant
    Dim i As Long

    ' Primary currency symbols to check
    symbols = Array(ChrW(163), "$", ChrW(8364))   ' GBP, USD, EUR
    symLabels = Array("GBP", "USD", "EUR")

    ' -- Check each symbol for format consistency --------------
    For i = LBound(symbols) To UBound(symbols)
        CheckSymbolConsistency doc, CStr(symbols(i)), CStr(symLabels(i)), issues
    Next i

    ' -- Check ISO code prefixed amounts -----------------------
    Dim isoCodes As Variant
    isoCodes = Array("GBP", "USD", "EUR", "JPY", "AUD", "CAD", "CHF", _
                     "BTC", "ETH", "USDT", "USDC", "BNB", "XRP", "SOL", "ADA", "DOGE")

    For i = LBound(isoCodes) To UBound(isoCodes)
        CheckISOCodeFormat doc, CStr(isoCodes(i)), issues
    Next i

    Set Check_CurrencyNumberFormat = issues
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.SetPageRange
' ----------------------------------------------------------------
Private Sub EngineSetPageRange(ByVal startPg As Long, ByVal endPg As Long)
    On Error Resume Next
    Application.Run "PleadingsEngine.SetPageRange", startPg, endPg
    If Err.Number <> 0 Then
        Debug.Print "EngineSetPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetDateFormatPref
' ----------------------------------------------------------------
Private Function EngineGetDateFormatPref() As String
    On Error Resume Next
    EngineGetDateFormatPref = Application.Run("PleadingsEngine.GetDateFormatPref")
    If Err.Number <> 0 Then
        Debug.Print "EngineGetDateFormatPref: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetDateFormatPref = "UK"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_Numbering.bas

```vb
Attribute VB_Name = "Rules_Numbering"
' ============================================================
' Rules_Numbering.bas
' Combined proofreading rules for numbering:
'   - Rule03: Sequential numbering (Check_SequentialNumbering)
'   - Rule08: Clause number format (Check_ClauseNumberFormat)
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME_SEQ As String = "sequential_numbering"
Private Const RULE_NAME_FMT As String = "clause_number_format"

' ============================================================
'  RULE 03 -- MAIN ENTRY POINT
' ============================================================
Public Function Check_SequentialNumbering(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Check Word-native numbered lists ------------------
    CheckNativeListNumbering doc, issues

    ' -- Check manually typed numbering --------------------
    CheckManualNumbering doc, issues

    Set Check_SequentialNumbering = issues
End Function

' ============================================================
'  PRIVATE: Check Word-native list numbering
'  Uses a Scripting.Dictionary keyed by list identifier to
'  track expected next values per list and level.
'
'  Each top-level key maps to a Dictionary of levels, where
'  each level stores the expected next value.
' ============================================================
Private Sub CheckNativeListNumbering(doc As Document, _
                                      ByRef issues As Collection)
    Dim listContexts As Object  ' listKey -> Dictionary(level -> expectedNext)
    Set listContexts = CreateObject("Scripting.Dictionary")
    Dim para As Paragraph
    Dim paraRange As Range
    Dim listType As Long
    Dim listKey As String
    Dim listLevel As Long
    Dim listValue As Long
    Dim expectedNext As Long
    Dim levelDict As Object
    Dim finding As Object
    Dim locStr As String
    Dim issueText As String
    Dim suggestion As String
    Dim prevLevel As Long

    ' Track the previous level per list to detect level changes
    Dim prevLevelDict As Object  ' listKey -> prevLevel
    Set prevLevelDict = CreateObject("Scripting.Dictionary")

    On Error Resume Next

    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextNativePara
        End If

        ' -- Skip non-list paragraphs ---------------------
        listType = paraRange.ListFormat.listType
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextNativePara
        End If

        ' wdListNoNumbering = 0; skip these
        If listType = 0 Then GoTo NextNativePara

        ' Only check numbered lists (wdListSimpleNumbering=1,
        ' wdListOutlineNumbering=4, wdListMixedNumbering=5)
        ' Skip bullet lists (wdListBullet=2, wdListPictureBullet=6)
        If listType = 2 Or listType = 6 Then GoTo NextNativePara

        ' -- Skip if outside configured page range --------
        If Not EngineIsInPageRange(paraRange) Then
            GoTo NextNativePara
        End If

        ' -- Determine list key (unique identifier) -------
        ' Try to use the List object's ListID first; fall back
        ' to a synthetic key built from type + position.
        listKey = ""
        Err.Clear
        Dim lstObj As Object
        Set lstObj = paraRange.ListFormat.List
        If Err.Number = 0 And Not lstObj Is Nothing Then
            listKey = "List_" & CStr(ObjPtr(lstObj))
        End If
        If Err.Number <> 0 Or Len(listKey) = 0 Then
            Err.Clear
            ' Synthetic key: use list type and an approximation
            listKey = "Synth_" & CStr(listType) & "_" & CStr(paraRange.ListFormat.ListLevelNumber)
        End If
        Err.Clear

        ' -- Get current list value and level -------------
        listValue = paraRange.ListFormat.listValue
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextNativePara
        End If

        listLevel = paraRange.ListFormat.ListLevelNumber
        If Err.Number <> 0 Then
            Err.Clear
            listLevel = 1
        End If

        ' -- Initialise tracking for this list if new -----
        If Not listContexts.Exists(listKey) Then
            Dim newLevelDict As Object
            Set newLevelDict = CreateObject("Scripting.Dictionary")
            listContexts.Add listKey, newLevelDict
            prevLevelDict.Add listKey, 0
        End If

        Set levelDict = listContexts(listKey)
        prevLevel = prevLevelDict(listKey)

        ' -- Handle level changes -------------------------
        ' When we go to a deeper level, that level starts fresh.
        ' When we return to a shallower level, reset all deeper levels.
        If listLevel <> prevLevel And prevLevel > 0 Then
            If listLevel < prevLevel Then
                ' Returning to shallower level: reset deeper levels
                Dim resetLevel As Variant
                Dim keysToRemove As New Collection
                For Each resetLevel In levelDict.keys
                    If CLng(resetLevel) > listLevel Then
                        keysToRemove.Add resetLevel
                    End If
                Next resetLevel
                Dim removeIdx As Long
                For removeIdx = 1 To keysToRemove.Count
                    levelDict.Remove keysToRemove(removeIdx)
                Next removeIdx
                Set keysToRemove = Nothing
            End If
        End If

        ' -- Check expected value at this level -----------
        If Not levelDict.Exists(listLevel) Then
            ' First item at this level in this list; record starting value
            levelDict.Add listLevel, listValue + 1
        Else
            expectedNext = levelDict(listLevel)

            If listValue = expectedNext Then
                ' Correct sequence; update expected next
                levelDict(listLevel) = listValue + 1

            ElseIf listValue = expectedNext - 1 Then
                ' Duplicate number
                Err.Clear
                locStr = EngineGetLocationString(paraRange, doc)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                End If

                issueText = "Duplicate number " & listValue & " at level " & listLevel
                suggestion = "Expected " & expectedNext & "; remove or renumber the duplicate"

                Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
                issues.Add finding
                ' Do not advance expectedNext for duplicates

            ElseIf listValue > expectedNext Then
                ' Skipped item(s)
                Err.Clear
                locStr = EngineGetLocationString(paraRange, doc)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                End If

                issueText = "Expected " & expectedNext & " but found " & listValue & _
                            " -- possible skipped item(s)"
                suggestion = "Check whether items " & expectedNext & " through " & _
                             (listValue - 1) & " are missing"

                Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
                issues.Add finding

                ' Update expected to continue from current
                levelDict(listLevel) = listValue + 1

            ElseIf listValue < expectedNext - 1 Then
                ' Numbering went backwards
                Err.Clear
                locStr = EngineGetLocationString(paraRange, doc)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                End If

                issueText = "Expected " & expectedNext & " but found " & listValue & _
                            " -- numbering went backwards"
                suggestion = "Renumber this item to " & expectedNext & " or check list continuity"

                Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
                issues.Add finding

                ' Update expected to continue from current
                levelDict(listLevel) = listValue + 1
            Else
                ' Normal sequence
                levelDict(listLevel) = listValue + 1
            End If
        End If

        ' Record previous level for this list
        prevLevelDict(listKey) = listLevel

NextNativePara:
    Next para
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Check manually typed numbering
'  Detects paragraphs that start with a number pattern
'  (e.g. "1.", "2.", "12.3") but have no Word list formatting.
'  Tracks these separately and checks for sequence breaks.
' ============================================================
Private Sub CheckManualNumbering(doc As Document, _
                                  ByRef issues As Collection)
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraText As String
    Dim listType As Long
    Dim manualNum As Long
    Dim expectedNext As Long
    Dim tracking As Boolean
    Dim finding As Object
    Dim locStr As String
    Dim issueText As String
    Dim suggestion As String
    Dim seqFontSize As Single   ' font size of the tracked sequence
    Dim seqIndent As Single     ' left indent of the tracked sequence
    Dim curFontSize As Single
    Dim curIndent As Single

    expectedNext = 0
    tracking = False
    seqFontSize = 0
    seqIndent = 0

    ' Track consecutive blank lines to detect section boundaries
    Dim blankLineRun As Long
    blankLineRun = 0

    On Error Resume Next

    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextManualPara
        End If

        paraText = Trim(paraRange.Text)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextManualPara
        End If

        ' -- Skip block quotes (they have their own numbering) --
        Dim isBlockQ As Boolean
        isBlockQ = False
        On Error Resume Next
        isBlockQ = Application.Run("Rules_Formatting.IsBlockQuotePara", para)
        If Err.Number <> 0 Then isBlockQ = False: Err.Clear
        On Error Resume Next
        If isBlockQ Then GoTo NextManualPara

        ' -- Only process non-list paragraphs -------------
        listType = paraRange.ListFormat.listType
        If Err.Number <> 0 Then
            Err.Clear
            listType = 0
        End If

        ' If this paragraph has Word list formatting, skip it
        ' and break any manual tracking chain
        If listType <> 0 Then
            tracking = False
            expectedNext = 0
            seqFontSize = 0
            seqIndent = 0
            blankLineRun = 0
            GoTo NextManualPara
        End If

        ' -- Check if paragraph starts with a number pattern -
        ' Patterns: "N." or "N)" where N is one or more digits
        manualNum = ExtractLeadingNumber(paraText)

        If manualNum < 0 Then
            ' No number pattern found
            If Len(paraText) <= 1 Then
                ' Blank/empty line: track consecutive blanks
                blankLineRun = blankLineRun + 1
                ' 3+ consecutive blank lines = likely section boundary
                If blankLineRun >= 3 And tracking Then
                    tracking = False
                    expectedNext = 0
                    seqFontSize = 0
                    seqIndent = 0
                End If
            Else
                blankLineRun = 0
                ' Check for section/schedule/annex headings that reset numbering
                Dim lcParaText As String
                lcParaText = LCase$(paraText)
                If lcParaText Like "schedule*" Or lcParaText Like "annex*" Or _
                   lcParaText Like "appendix*" Or lcParaText Like "part *" Or _
                   lcParaText Like "section *" Then
                    tracking = False
                    expectedNext = 0
                    seqFontSize = 0
                    seqIndent = 0
                Else
                    ' Substantial non-numbered text: break tracking chain
                    tracking = False
                    expectedNext = 0
                    seqFontSize = 0
                    seqIndent = 0
                End If
            End If
            GoTo NextManualPara
        End If

        blankLineRun = 0

        ' -- Skip if outside configured page range --------
        If Not EngineIsInPageRange(paraRange) Then
            GoTo NextManualPara
        End If

        ' -- Detect context change (block quote / different list) --
        ' If this numbered paragraph has a significantly different
        ' font size or indentation from the current sequence, it
        ' belongs to a different numbering context (e.g. a block
        ' quote with its own numbering). Skip it without breaking
        ' the tracking chain.
        curFontSize = 0
        curFontSize = paraRange.Font.Size
        If Err.Number <> 0 Then curFontSize = 0: Err.Clear
        curIndent = 0
        curIndent = para.Format.LeftIndent
        If Err.Number <> 0 Then curIndent = 0: Err.Clear

        If tracking And curFontSize > 0 And seqFontSize > 0 Then
            ' Font size differs by more than 1pt = different context
            If Abs(curFontSize - seqFontSize) > 1 Then GoTo NextManualPara
            ' Indentation differs by more than 36pt (0.5 inch) = different context
            If Abs(curIndent - seqIndent) > 36 Then GoTo NextManualPara
        End If

        ' -- Start or continue tracking -------------------
        If Not tracking Then
            ' First manually numbered paragraph in a sequence
            tracking = True
            expectedNext = manualNum + 1
            seqFontSize = curFontSize
            seqIndent = curIndent
            GoTo NextManualPara
        End If

        ' -- Check sequence -------------------------------
        If manualNum = expectedNext Then
            ' Correct sequence
            expectedNext = manualNum + 1

        ElseIf manualNum > expectedNext Then
            ' Skipped item(s)
            Err.Clear
            locStr = EngineGetLocationString(paraRange, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If

            issueText = "Manual numbering: expected " & expectedNext & _
                        " but found " & manualNum & " -- possible skipped item(s)"
            suggestion = "Check whether items " & expectedNext & " through " & _
                         (manualNum - 1) & " are missing"

            Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
            issues.Add finding

            expectedNext = manualNum + 1

        ElseIf manualNum < expectedNext And manualNum = expectedNext - 1 Then
            ' Duplicate
            Err.Clear
            locStr = EngineGetLocationString(paraRange, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If

            issueText = "Manual numbering: duplicate number " & manualNum
            suggestion = "Remove or renumber the duplicate item"

            Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
            issues.Add finding

        ElseIf manualNum < expectedNext - 1 Then
            ' Backwards
            Err.Clear
            locStr = EngineGetLocationString(paraRange, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If

            issueText = "Manual numbering: expected " & expectedNext & _
                        " but found " & manualNum & " -- numbering went backwards"
            suggestion = "Renumber this item to " & expectedNext & " or check sequence"

            Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
            issues.Add finding

            expectedNext = manualNum + 1
        Else
            ' Normal (covers any other case)
            expectedNext = manualNum + 1
        End If

NextManualPara:
    Next para
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Extract leading number from paragraph text
'  Returns the number if the text starts with a pattern like
'  "1.", "12.", "3)", "42)"; returns -1 if no match.
'  Uses the VBA Like operator for pattern matching.
' ============================================================
Private Function ExtractLeadingNumber(ByVal txt As String) As Long
    Dim trimmed As String
    Dim numStr As String
    Dim i As Long
    Dim ch As String

    trimmed = Trim(txt)
    ExtractLeadingNumber = -1

    If Len(trimmed) = 0 Then Exit Function

    ' Check first character is a digit
    If Not (trimmed Like "#*") Then Exit Function

    ' Extract consecutive digits from the start
    numStr = ""
    For i = 1 To Len(trimmed)
        ch = Mid(trimmed, i, 1)
        If ch >= "0" And ch <= "9" Then
            numStr = numStr & ch
        Else
            Exit For
        End If
    Next i

    ' Must have at least one digit
    If Len(numStr) = 0 Then Exit Function

    ' The character after the digits must be "." or ")"
    ' to qualify as a numbering pattern
    If i <= Len(trimmed) Then
        ch = Mid(trimmed, i, 1)
        If ch = "." Or ch = ")" Then
            On Error Resume Next
            ExtractLeadingNumber = CLng(numStr)
            If Err.Number <> 0 Then
                ExtractLeadingNumber = -1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If
End Function

' ============================================================
'  PRIVATE: Extract clause number prefix from paragraph text
'  Returns the clause number prefix or empty string if none
'  found.
' ============================================================
Private Function ExtractClausePrefix(ByVal txt As String) As String
    Dim cleanText As String
    cleanText = Trim$(Replace(txt, vbCr, ""))
    cleanText = Trim$(Replace(cleanText, vbLf, ""))
    If Len(cleanText) = 0 Then
        ExtractClausePrefix = ""
        Exit Function
    End If

    ' A clause number starts at the beginning and ends before
    ' the first space or tab that is followed by non-number text
    Dim i As Long
    Dim ch As String
    Dim prefix As String
    prefix = ""

    ' Must start with a digit
    If Not (Left$(cleanText, 1) Like "[0-9]") Then
        ExtractClausePrefix = ""
        Exit Function
    End If

    ' Collect characters that form the clause number
    ' Valid clause number chars: digits, dots, parens, lowercase letters
    For i = 1 To Len(cleanText)
        ch = Mid$(cleanText, i, 1)
        If ch Like "[0-9]" Or ch = "." Or ch = "(" Or ch = ")" Or _
           (ch Like "[a-z]" And i > 1) Or (ch Like "[ivxlcdm]" And i > 1) Then
            prefix = prefix & ch
        ElseIf ch = " " Or ch = vbTab Or ch = Chr(9) Then
            ' End of clause number
            Exit For
        Else
            ' Non-clause character encountered
            Exit For
        End If
    Next i

    ' Validate: must contain at least one digit
    Dim hasDigit As Boolean
    hasDigit = False
    For i = 1 To Len(prefix)
        If Mid$(prefix, i, 1) Like "[0-9]" Then
            hasDigit = True
            Exit For
        End If
    Next i

    If Not hasDigit Then
        ExtractClausePrefix = ""
        Exit Function
    End If

    ' Remove trailing dots (e.g., "1." -> "1")
    Do While Len(prefix) > 0 And Right$(prefix, 1) = "."
        prefix = Left$(prefix, Len(prefix) - 1)
    Loop

    ExtractClausePrefix = prefix
End Function

' ============================================================
'  PRIVATE: Classify the clause number format
'  Returns a format pattern string describing the style
' ============================================================
Private Function ClassifyClauseFormat(ByVal prefix As String) As String
    If Len(prefix) = 0 Then
        ClassifyClauseFormat = ""
        Exit Function
    End If

    ' Level 1: plain number like "1" or "12"
    If prefix Like "#" Or prefix Like "##" Or prefix Like "###" Then
        ClassifyClauseFormat = "L1_plain"
        Exit Function
    End If

    ' Level 2: dotted like "1.1", "12.3", "1.12"
    If prefix Like "#.#" Or prefix Like "##.#" Or prefix Like "#.##" Or _
       prefix Like "##.##" Then
        ClassifyClauseFormat = "L2_dotted"
        Exit Function
    End If

    ' Level 3 style A: "1.1(a)" -- dotted number followed by (letter)
    If prefix Like "#.#(*)" Or prefix Like "##.#(*)" Or _
       prefix Like "#.##(*)" Or prefix Like "##.##(*)" Then
        ' Check if content in parens is a lowercase letter
        Dim parenContent As String
        Dim pOpen As Long
        pOpen = InStr(1, prefix, "(")
        If pOpen > 0 Then
            Dim pClose As Long
            pClose = InStr(pOpen, prefix, ")")
            If pClose > pOpen + 1 Then
                parenContent = Mid$(prefix, pOpen + 1, pClose - pOpen - 1)
                If Len(parenContent) = 1 And parenContent Like "[a-z]" Then
                    ClassifyClauseFormat = "L3_dotted_letter"
                    Exit Function
                End If
            End If
        End If
    End If

    ' Level 3 style B: "1.1.1" -- triple dotted
    If prefix Like "#.#.#" Or prefix Like "##.#.#" Or _
       prefix Like "#.##.#" Or prefix Like "#.#.##" Then
        ClassifyClauseFormat = "L3_dotted_sub"
        Exit Function
    End If

    ' Level 4: double parenthetical like "1.1(a)(i)"
    Dim parenCount As Long
    Dim ci As Long
    parenCount = 0
    For ci = 1 To Len(prefix)
        If Mid$(prefix, ci, 1) = "(" Then parenCount = parenCount + 1
    Next ci
    If parenCount >= 2 Then
        ClassifyClauseFormat = "L4_double_paren"
        Exit Function
    End If

    ' Single parenthetical at end: "(a)" or "(i)" style
    If Right$(prefix, 1) = ")" Then
        pOpen = InStrRev(prefix, "(")
        If pOpen > 0 Then
            parenContent = Mid$(prefix, pOpen + 1, Len(prefix) - pOpen - 1)
            If Len(parenContent) = 1 And parenContent Like "[a-z]" Then
                ClassifyClauseFormat = "L3_paren_letter"
                Exit Function
            End If
            ' Roman numeral in parens
            Dim allRoman As Boolean
            allRoman = True
            For ci = 1 To Len(parenContent)
                If Not (Mid$(parenContent, ci, 1) Like "[ivxlcdm]") Then
                    allRoman = False
                    Exit For
                End If
            Next ci
            If allRoman And Len(parenContent) > 0 Then
                ClassifyClauseFormat = "L3_paren_roman"
                Exit Function
            End If
        End If
    End If

    ' Fallback: generic numbered
    ClassifyClauseFormat = "other_" & prefix
End Function

' ============================================================
'  RULE 08 -- MAIN ENTRY POINT
' ============================================================
Public Function Check_ClauseNumberFormat(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long

    On Error Resume Next

    ' Track format patterns: formatPattern -> Collection of Array(paraIdx, prefix, rangeStart, rangeEnd)
    Dim formatCounts As Object
    Set formatCounts = CreateObject("Scripting.Dictionary")
    Dim clauseInfos As New Collection
    Dim cInfo() As Variant

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        ' Skip headings (they have their own numbering rules)
        If para.OutlineLevel >= wdOutlineLevel1 And _
           para.OutlineLevel <= wdOutlineLevel9 Then GoTo NextClausePara

        ' Page range filter
        If Not EngineIsInPageRange(para.Range) Then GoTo NextClausePara

        ' Extract clause number prefix
        Dim prefix As String
        prefix = ExtractClausePrefix(para.Range.Text)
        If Len(prefix) = 0 Then GoTo NextClausePara

        ' Classify the format
        Dim fmt As String
        fmt = ClassifyClauseFormat(prefix)
        If Len(fmt) = 0 Then GoTo NextClausePara

        ' Count format occurrences
        If formatCounts.Exists(fmt) Then
            formatCounts(fmt) = formatCounts(fmt) + 1
        Else
            formatCounts.Add fmt, 1
        End If

        ' Store clause info
        ReDim cInfo(0 To 3)
        cInfo(0) = paraIdx
        cInfo(1) = prefix
        cInfo(2) = para.Range.Start
        cInfo(3) = para.Range.End
        clauseInfos.Add Array(fmt, cInfo)
NextClausePara:
    Next para

    ' -- Group into contiguous runs for local-context comparison --
    ' A new run starts when paragraph gap > MAX_CLAUSE_GAP.
    ' Within each run, compare formats per level category.
    Const MAX_CLAUSE_GAP As Long = 40
    If clauseInfos.Count < 2 Then GoTo ClauseFormatDone

    ' Build arrays for grouping
    Dim ciTotal As Long
    ciTotal = clauseInfos.Count
    Dim ciFmts() As String, ciIdxs() As Long
    Dim ciStarts() As Long, ciEnds() As Long
    Dim ciPrefixes() As String
    ReDim ciFmts(0 To ciTotal - 1)
    ReDim ciIdxs(0 To ciTotal - 1)
    ReDim ciStarts(0 To ciTotal - 1)
    ReDim ciEnds(0 To ciTotal - 1)
    ReDim ciPrefixes(0 To ciTotal - 1)

    Dim ci As Long
    For ci = 1 To ciTotal
        Dim clauseArr As Variant
        clauseArr = clauseInfos(ci)
        Dim clauseData As Variant
        clauseData = clauseArr(1)
        ciFmts(ci - 1) = CStr(clauseArr(0))
        ciIdxs(ci - 1) = CLng(clauseData(0))
        ciStarts(ci - 1) = CLng(clauseData(2))
        ciEnds(ci - 1) = CLng(clauseData(3))
        ciPrefixes(ci - 1) = CStr(clauseData(1))
    Next ci

    ' Identify run boundaries
    Dim runStart As Long, ri As Long
    runStart = 0

    For ri = 0 To ciTotal  ' one past end to close final run
        Dim startNewRun As Boolean
        startNewRun = (ri = ciTotal)  ' always close at end
        If Not startNewRun And ri > 0 Then
            If ciIdxs(ri) - ciIdxs(ri - 1) > MAX_CLAUSE_GAP Then
                startNewRun = True
            End If
        End If

        If startNewRun And ri > runStart Then
            ' Process the run [runStart..ri-1]
            Dim runEnd As Long
            runEnd = ri - 1
            If runEnd > runStart Then  ' need >= 2 items
                ' Group by level category within this run
                Dim runLevelGroups As Object
                Set runLevelGroups = CreateObject("Scripting.Dictionary")
                Dim rj As Long
                For rj = runStart To runEnd
                    Dim levelCat As String
                    If Left$(ciFmts(rj), 2) = "L1" Then
                        levelCat = "L1"
                    ElseIf Left$(ciFmts(rj), 2) = "L2" Then
                        levelCat = "L2"
                    ElseIf Left$(ciFmts(rj), 2) = "L3" Then
                        levelCat = "L3"
                    ElseIf Left$(ciFmts(rj), 2) = "L4" Then
                        levelCat = "L4"
                    Else
                        levelCat = "other"
                    End If

                    If Not runLevelGroups.Exists(levelCat) Then
                        runLevelGroups.Add levelCat, CreateObject("Scripting.Dictionary")
                    End If
                    Dim lgDict As Object
                    Set lgDict = runLevelGroups(levelCat)
                    If lgDict.Exists(ciFmts(rj)) Then
                        lgDict(ciFmts(rj)) = lgDict(ciFmts(rj)) + 1
                    Else
                        lgDict.Add ciFmts(rj), 1
                    End If
                Next rj

                ' Find dominant per level in this run
                Dim lgKey As Variant
                Dim fk As Variant
                For Each lgKey In runLevelGroups.keys
                    Set lgDict = runLevelGroups(lgKey)
                    If lgDict.Count > 1 Then
                        Dim domFmt As String
                        Dim maxCnt As Long
                        domFmt = ""
                        maxCnt = 0
                        For Each fk In lgDict.keys
                            If lgDict(fk) > maxCnt Then
                                maxCnt = lgDict(fk)
                                domFmt = CStr(fk)
                            End If
                        Next fk

                        ' Flag deviations within this run
                        For rj = runStart To runEnd
                            Dim rjLevelCat As String
                            If Left$(ciFmts(rj), 2) = "L1" Then
                                rjLevelCat = "L1"
                            ElseIf Left$(ciFmts(rj), 2) = "L2" Then
                                rjLevelCat = "L2"
                            ElseIf Left$(ciFmts(rj), 2) = "L3" Then
                                rjLevelCat = "L3"
                            ElseIf Left$(ciFmts(rj), 2) = "L4" Then
                                rjLevelCat = "L4"
                            Else
                                rjLevelCat = "other"
                            End If

                            If rjLevelCat = CStr(lgKey) And ciFmts(rj) <> domFmt Then
                                Dim finding As Object
                                Dim rng As Range
                                Set rng = doc.Range(ciStarts(rj), ciEnds(rj))
                                Dim loc As String
                                loc = EngineGetLocationString(rng, doc)
                                If Err.Number <> 0 Then loc = "unknown location": Err.Clear

                                Set finding = CreateIssueDict(RULE_NAME_FMT, loc, _
                                    "Mixed clause number format: '" & ciPrefixes(rj) & _
                                    "' uses style " & ciFmts(rj) & " but dominant " & _
                                    CStr(lgKey) & " style in this section is " & domFmt, _
                                    "Reformat to match the dominant clause numbering style", _
                                    ciStarts(rj), ciEnds(rj), "error")
                                issues.Add finding
                            End If
                        Next rj
                    End If
                Next lgKey
            End If
            runStart = ri
        End If
    Next ri
ClauseFormatDone:

    On Error GoTo 0
    Set Check_ClauseNumberFormat = issues
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_Punctuation.bas

```vb
Attribute VB_Name = "Rules_Punctuation"
' ============================================================
' Rules_Punctuation.bas
' Combined proofreading rules for punctuation checking:
'   - Slash style (Rule14): checks forward slash spacing
'     consistency and flags unexpected backslashes.
'   - Bracket integrity (Rule16): checks for mismatched,
'     unmatched, and improperly nested brackets: (), [], {}.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME_SLASH As String = "slash_style"
Private Const RULE_NAME_BRACKET As String = "bracket_integrity"
Private Const RULE_NAME_DASH As String = "dash_usage"

' ?==============================================================?
' ?  SLASH STYLE (Rule14)                                       ?
' ?==============================================================?

' ============================================================
'  MAIN ENTRY POINT: Slash Style
' ============================================================
Public Function Check_SlashStyle(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Forward slashes: determine dominant style ------------
    Dim tightCount As Long
    Dim spacedCount As Long

    tightCount = CountTightSlashes(doc)
    spacedCount = CountSpacedSlashes(doc)

    ' Determine dominant style
    Dim dominantStyle As String
    If tightCount >= spacedCount Then
        dominantStyle = "tight"
    Else
        dominantStyle = "spaced"
    End If

    ' Flag minority style forward slashes
    If dominantStyle = "tight" And spacedCount > 0 Then
        FlagSpacedSlashes doc, issues
    ElseIf dominantStyle = "spaced" And tightCount > 0 Then
        FlagTightSlashes doc, issues
    End If

    ' -- Backslashes ------------------------------------------
    FlagBackslashes doc, issues

    Set Check_SlashStyle = issues
End Function

' ============================================================
'  PRIVATE: Count tight slashes using wildcard search
'  Excludes conventional tight pairs (and/or, his/her, etc.)
'  so they don't bias the dominant-style determination.
' ============================================================
Private Function CountTightSlashes(doc As Document) As Long
    Dim rng As Range
    Dim cnt As Long
    Dim found As Boolean

    cnt = 0
    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = "[! ]/[! ]"
        .MatchWildcards = True
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        ' Skip URLs and dates
        If Not IsURLContext(rng, doc) And Not IsDateSlash(rng) Then
            ' Skip conventional tight pairs (and/or, his/her, etc.)
            If Not IsConventionalTightSlash(rng, doc) Then
                cnt = cnt + 1
            End If
        End If

        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0

    CountTightSlashes = cnt
End Function

' ============================================================
'  PRIVATE: Count spaced slashes using literal search
' ============================================================
Private Function CountSpacedSlashes(doc As Document) As Long
    Dim rng As Range
    Dim cnt As Long
    Dim found As Boolean

    cnt = 0
    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = " / "
        .MatchWildcards = False
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        ' Skip URLs
        If Not IsURLContext(rng, doc) Then
            cnt = cnt + 1
        End If

        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0

    CountSpacedSlashes = cnt
End Function

' ============================================================
'  PRIVATE: Flag spaced slashes (minority when tight is dominant)
' ============================================================
Private Sub FlagSpacedSlashes(doc As Document, ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String

    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = " / "
        .MatchWildcards = False
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        If Not EngineIsInPageRange(rng) Then GoTo ContinueSpaced
        If IsURLContext(rng, doc) Then GoTo ContinueSpaced

        locStr = EngineGetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set finding = CreateIssueDict(RULE_NAME_SLASH, locStr, "Spaced slash '" & rng.Text & "' differs from dominant tight style", "Remove spaces around slash for consistency", rng.Start, rng.End, "possible_error")
        issues.Add finding

ContinueSpaced:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Flag tight slashes (minority when spaced is dominant)
' ============================================================
Private Sub FlagTightSlashes(doc As Document, ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String

    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = "[! ]/[! ]"
        .MatchWildcards = True
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos2 As Long
    lastPos2 = -1
    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do
        If rng.Start <= lastPos2 Then Exit Do   ' stall guard
        lastPos2 = rng.Start

        If Not EngineIsInPageRange(rng) Then GoTo ContinueTight
        If IsURLContext(rng, doc) Then GoTo ContinueTight
        If IsDateSlash(rng) Then GoTo ContinueTight
        If IsConventionalTightSlash(rng, doc) Then GoTo ContinueTight

        locStr = EngineGetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set finding = CreateIssueDict(RULE_NAME_SLASH, locStr, "Tight slash '" & rng.Text & "' differs from dominant spaced style", "Add spaces around slash for consistency", rng.Start, rng.End, "possible_error")
        issues.Add finding

ContinueTight:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Flag unexpected backslashes
' ============================================================
Private Sub FlagBackslashes(doc As Document, ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String
    Dim context As String
    Dim fontName As String

    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = "\"
        .MatchWildcards = False
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos3 As Long
    lastPos3 = -1
    On Error Resume Next
    Do
        Err.Clear
        found = rng.Find.Execute
        If Err.Number <> 0 Then Exit Do
        If Not found Then Exit Do
        If rng.Start <= lastPos3 Then Exit Do   ' stall guard
        lastPos3 = rng.Start

        If Not EngineIsInPageRange(rng) Then GoTo ContinueBackslash

        ' Get surrounding context for skip checks
        Dim contextStart As Long
        Dim contextEnd As Long
        Dim contextRng As Range

        contextStart = rng.Start - 5
        If contextStart < 0 Then contextStart = 0
        contextEnd = rng.End + 10
        If contextEnd > doc.Content.End Then contextEnd = doc.Content.End

        Set contextRng = doc.Range(contextStart, contextEnd)
        If Err.Number <> 0 Then
            Err.Clear
            context = ""
        Else
            context = LCase(contextRng.Text)
        End If

        ' Skip file paths: drive letter pattern like C:\ or UNC \\server
        If IsDriveLetterPath(context) Or IsUNCPath(context) Then
            GoTo ContinueBackslash
        End If

        ' Skip code-font runs (Courier, Consolas)
        fontName = ""
        fontName = rng.Font.Name
        If Err.Number <> 0 Then
            Err.Clear
            fontName = ""
        End If
        If IsCodeFontName(fontName) Then
            GoTo ContinueBackslash
        End If

        ' Skip URLs
        If InStr(1, context, "://") > 0 Then
            GoTo ContinueBackslash
        End If

        ' Flag the backslash
        locStr = EngineGetLocationString(rng, doc)
        If Err.Number <> 0 Then
            locStr = "unknown location"
            Err.Clear
        End If

        Set finding = CreateIssueDict(RULE_NAME_SLASH, locStr, "Unexpected backslash -- did you mean forward slash?", "Replace '\' with '/'", rng.Start, rng.End, "possible_error")
        issues.Add finding

ContinueBackslash:
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Check if a tight slash is a conventional pair
'  (and/or, his/her, etc.) that should always be tight
'  regardless of the document's dominant slash style.
' ============================================================
Private Function IsConventionalTightSlash(rng As Range, doc As Document) As Boolean
    IsConventionalTightSlash = False
    On Error Resume Next

    ' Expand range to capture surrounding word context
    Dim ctxStart As Long, ctxEnd As Long
    ctxStart = rng.Start - 12
    If ctxStart < 0 Then ctxStart = 0
    ctxEnd = rng.End + 12
    If ctxEnd > doc.Content.End Then ctxEnd = doc.Content.End

    Dim ctxRng As Range
    Set ctxRng = doc.Range(ctxStart, ctxEnd)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function

    Dim ctxText As String
    ctxText = LCase$(ctxRng.Text)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    On Error GoTo 0

    ' Known conventionally-tight slash pairs
    Dim tightPairs As Variant
    tightPairs = Array("and/or", "either/or", "his/her", "he/she", _
                       "s/he", "w/o", "n/a", "c/o", "a/c", "y/n", "yes/no", _
                       "input/output", "read/write", "true/false", "on/off", _
                       "open/close", "start/end", "pass/fail", "client/server")
    Dim tp As Variant
    For Each tp In tightPairs
        If InStr(1, ctxText, CStr(tp), vbTextCompare) > 0 Then
            IsConventionalTightSlash = True
            Exit Function
        End If
    Next tp

    ' General word/word alternative detection:
    ' If the match text is letter(s)/letter(s) and both sides are short
    ' English words (2-12 chars), treat as a functional alternative pair.
    Dim matchText As String
    matchText = LCase$(rng.Text)
    Dim slashPos As Long
    slashPos = InStr(1, matchText, "/")
    If slashPos > 1 And slashPos < Len(matchText) Then
        Dim lWord As String, rWord As String
        lWord = Left$(matchText, slashPos - 1)
        rWord = Mid$(matchText, slashPos + 1)
        ' Both sides are purely alphabetic and short
        If Len(lWord) >= 2 And Len(lWord) <= 12 And _
           Len(rWord) >= 2 And Len(rWord) <= 12 Then
            If IsAlphaOnly(lWord) And IsAlphaOnly(rWord) Then
                IsConventionalTightSlash = True
                Exit Function
            End If
        End If
    End If
End Function

' Helper: check if a string is purely alphabetic (a-z)
Private Function IsAlphaOnly(ByVal s As String) As Boolean
    Dim ci As Long
    For ci = 1 To Len(s)
        Dim cc As String
        cc = Mid$(s, ci, 1)
        If Not ((cc >= "a" And cc <= "z") Or (cc >= "A" And cc <= "Z")) Then
            IsAlphaOnly = False
            Exit Function
        End If
    Next ci
    IsAlphaOnly = True
End Function

' ============================================================
'  PRIVATE: Check if context suggests a URL
' ============================================================
Private Function IsURLContext(rng As Range, doc As Document) As Boolean
    Dim contextStart As Long
    Dim contextEnd As Long
    Dim contextRng As Range
    Dim context As String

    On Error Resume Next
    contextStart = rng.Start - 30
    If contextStart < 0 Then contextStart = 0
    contextEnd = rng.End + 30
    If contextEnd > doc.Content.End Then contextEnd = doc.Content.End

    Set contextRng = doc.Range(contextStart, contextEnd)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsURLContext = False
        Exit Function
    End If

    context = LCase(contextRng.Text)
    On Error GoTo 0

    IsURLContext = (InStr(1, context, "://") > 0) Or _
                   (InStr(1, context, "http") > 0) Or _
                   (InStr(1, context, "www") > 0)
End Function

' ============================================================
'  PRIVATE: Check if slash is part of a date (digits only)
' ============================================================
Private Function IsDateSlash(rng As Range) As Boolean
    Dim matchText As String
    Dim i As Long
    Dim ch As String
    Dim hasSlash As Boolean

    matchText = rng.Text
    If Len(matchText) < 3 Then
        IsDateSlash = False
        Exit Function
    End If

    ' Check that all non-slash characters are digits
    hasSlash = False
    For i = 1 To Len(matchText)
        ch = Mid(matchText, i, 1)
        If ch = "/" Then
            hasSlash = True
        ElseIf Not (ch >= "0" And ch <= "9") Then
            IsDateSlash = False
            Exit Function
        End If
    Next i

    IsDateSlash = hasSlash
End Function

' ============================================================
'  PRIVATE: Check for drive letter path pattern (e.g. C:\)
' ============================================================
Private Function IsDriveLetterPath(ByVal context As String) As Boolean
    Dim i As Long
    Dim ch As String

    ' Look for pattern: letter followed by :\
    For i = 1 To Len(context) - 2
        ch = Mid(context, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Then
            If Mid(context, i + 1, 2) = ":\" Then
                IsDriveLetterPath = True
                Exit Function
            End If
        End If
    Next i

    IsDriveLetterPath = False
End Function

' ============================================================
'  PRIVATE: Check for UNC path pattern (\\server)
' ============================================================
Private Function IsUNCPath(ByVal context As String) As Boolean
    IsUNCPath = (InStr(1, context, "\\") > 0)
End Function

' ?==============================================================?
' ?  BRACKET INTEGRITY (Rule16)                                 ?
' ?==============================================================?

' ============================================================
'  MAIN ENTRY POINT: Bracket Integrity
' ============================================================
Public Function Check_BracketIntegrity(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraText As String
    Dim paraStart As Long

    ' Counters per bracket type (reset per paragraph)
    Dim parenOpen As Long, parenClose As Long
    Dim sqOpen As Long, sqClose As Long
    Dim curlyOpen As Long, curlyClose As Long

    ' Position of first unmatched bracket (for issue location)
    Dim firstParenPos As Long, firstSqPos As Long, firstCurlyPos As Long

    Dim b() As Byte, bMax As Long
    Dim i As Long, code As Long, pos As Long

    For Each para In doc.Paragraphs
        On Error Resume Next
        paraText = para.Range.Text
        paraStart = para.Range.Start
        If Err.Number <> 0 Then
            Err.Clear: On Error GoTo 0
            GoTo NxtPara
        End If
        On Error GoTo 0

        If LenB(paraText) = 0 Then GoTo NxtPara

        ' Compute list prefix length for position correction
        Dim bktListPrefixLen As Long
        bktListPrefixLen = GetDashListPrefixLen(para, paraText)

        ' Reset counters
        parenOpen = 0: parenClose = 0
        sqOpen = 0: sqClose = 0
        curlyOpen = 0: curlyClose = 0
        firstParenPos = -1: firstSqPos = -1: firstCurlyPos = -1

        b = paraText
        bMax = UBound(b) - 1

        For i = 0 To bMax Step 2
            code = b(i) Or (CLng(b(i + 1)) * 256&)
            pos = paraStart + (i \ 2) - bktListPrefixLen

            Select Case code
                Case 40   ' (
                    parenOpen = parenOpen + 1
                    If firstParenPos < 0 Then firstParenPos = pos
                Case 41   ' )
                    parenClose = parenClose + 1
                    If firstParenPos < 0 Then firstParenPos = pos
                Case 91   ' [
                    sqOpen = sqOpen + 1
                    If firstSqPos < 0 Then firstSqPos = pos
                Case 93   ' ]
                    sqClose = sqClose + 1
                    If firstSqPos < 0 Then firstSqPos = pos
                Case 123  ' {
                    curlyOpen = curlyOpen + 1
                    If firstCurlyPos < 0 Then firstCurlyPos = pos
                Case 125  ' }
                    curlyClose = curlyClose + 1
                    If firstCurlyPos < 0 Then firstCurlyPos = pos
            End Select
        Next i

        ' Report once per bracket type if counts don't match
        If parenOpen <> parenClose Then
            CreateBracketIssue doc, issues, firstParenPos, "()", _
                "Unbalanced parentheses: " & parenOpen & " opened, " & _
                parenClose & " closed"
        End If
        If sqOpen <> sqClose Then
            CreateBracketIssue doc, issues, firstSqPos, "[]", _
                "Unbalanced square brackets: " & sqOpen & " opened, " & _
                sqClose & " closed"
        End If
        If curlyOpen <> curlyClose Then
            CreateBracketIssue doc, issues, firstCurlyPos, "{}", _
                "Unbalanced curly braces: " & curlyOpen & " opened, " & _
                curlyClose & " closed"
        End If

        ' -- Stack-based nesting check (only when counts balance) --
        If parenOpen = parenClose And sqOpen = sqClose _
           And curlyOpen = curlyClose _
           And (parenOpen + sqOpen + curlyOpen) > 0 Then
            Dim stk() As Long, stkTop As Long
            stkTop = 0
            ReDim stk(1 To parenOpen + sqOpen + curlyOpen)
            Dim nestBad As Boolean, nestPos As Long
            nestBad = False
            For i = 0 To bMax Step 2
                code = b(i) Or (CLng(b(i + 1)) * 256&)
                Select Case code
                    Case 40, 91, 123  ' open bracket
                        stkTop = stkTop + 1
                        If stkTop > UBound(stk) Then ReDim Preserve stk(1 To stkTop + 4)
                        stk(stkTop) = code
                    Case 41, 93, 125  ' close bracket
                        If stkTop = 0 Then
                            nestBad = True
                            nestPos = paraStart + (i \ 2) - bktListPrefixLen
                            Exit For
                        End If
                        If Not CodesMatch(stk(stkTop), code) Then
                            nestBad = True
                            nestPos = paraStart + (i \ 2) - bktListPrefixLen
                            Exit For
                        End If
                        stkTop = stkTop - 1
                End Select
            Next i
            If nestBad Then
                CreateBracketIssue doc, issues, nestPos, "()", _
                    "Improperly nested brackets (e.g. overlapping pairs)"
            End If
        End If

NxtPara:
    Next para

    Set Check_BracketIntegrity = issues
End Function

' ------------------------------------------------------------
'  Code-point bracket matching (no string comparison)
' ------------------------------------------------------------
Private Function CodesMatch(ByVal openCode As Long, _
        ByVal closeCode As Long) As Boolean
    Select Case openCode
        Case 40:  CodesMatch = (closeCode = 41)   ' ( -> )
        Case 91:  CodesMatch = (closeCode = 93)   ' [ -> ]
        Case 123: CodesMatch = (closeCode = 125)  ' { -> }
        Case Else: CodesMatch = False
    End Select
End Function

' ============================================================
'  PRIVATE: Create a bracket integrity finding
' ============================================================
Private Sub CreateBracketIssue(doc As Document, _
                                ByRef issues As Collection, _
                                ByVal pos As Long, _
                                ByVal bracketChar As String, _
                                ByVal issueText As String)
    Dim finding As Object
    Dim locStr As String
    Dim rng As Range

    On Error Resume Next
    Set rng = doc.Range(pos, pos + 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    ' Skip if outside page range
    If Not EngineIsInPageRange(rng) Then
        On Error GoTo 0
        Exit Sub
    End If

    locStr = EngineGetLocationString(rng, doc)
    If Err.Number <> 0 Then
        locStr = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0

    ' Determine suggestion based on bracket type
    Dim suggestion As String
    Select Case bracketChar
        Case "()", "(", ")"
            suggestion = "Add or correct matching parenthesis"
        Case "[]", "[", "]"
            suggestion = "Add or correct matching square bracket"
        Case "{}", "{", "}"
            suggestion = "Add or correct matching curly brace"
        Case Else
            suggestion = "Review bracket pairing"
    End Select

    Set finding = CreateIssueDict(RULE_NAME_BRACKET, locStr, issueText, suggestion, pos, pos + 1, "error")
    issues.Add finding
End Sub

' ?==============================================================?
' ?  SHARED PRIVATE HELPERS                                     ?
' ?==============================================================?

' ============================================================
'  PRIVATE: Check if a font name is a code font (Courier, Consolas)
'  Shared by FlagBackslashes and IsCodeFont
' ============================================================
Private Function IsCodeFontName(ByVal fontName As String) As Boolean
    IsCodeFontName = (LCase(fontName) Like "*courier*") Or _
                     (LCase(fontName) Like "*consolas*")
End Function


' Calculate the length of auto-generated list numbering text
Private Function GetDashListPrefixLen(para As Paragraph, ByVal paraText As String) As Long
    GetDashListPrefixLen = 0
    On Error Resume Next
    Dim lStr As String
    lStr = para.Range.ListFormat.ListString
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    If Len(lStr) = 0 Then On Error GoTo 0: Exit Function
    If Len(paraText) > Len(lStr) Then
        If Left$(paraText, Len(lStr)) = lStr Then
            GetDashListPrefixLen = Len(lStr)
            If Mid$(paraText, GetDashListPrefixLen + 1, 1) = vbTab Then
                GetDashListPrefixLen = GetDashListPrefixLen + 1
            End If
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ================================================================
' ================================================================
'  DASH USAGE (en-dash / em-dash / hyphen)
'  Context-dependent checks:
'   1. Hyphen in number ranges -> should be en-dash
'   2. Double-hyphen "--" -> should be em-dash
'   3. En-dash between words (compound) -> should be hyphen
'   4. Spaced en-dash -> should probably be em-dash
' ================================================================
' ================================================================

Public Function Check_DashUsage(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraText As String
    Dim paraRange As Range
    Dim finding As Object
    Dim locStr As String

    Dim reHyphenRange As Object
    Set reHyphenRange = CreateObject("VBScript.RegExp")
    reHyphenRange.Global = True
    ' Matches digit(s) - hyphen - digit(s) as a number range
    reHyphenRange.Pattern = "(\d)-(\d)"

    Dim reDoubleHyphen As Object
    Set reDoubleHyphen = CreateObject("VBScript.RegExp")
    reDoubleHyphen.Global = True
    reDoubleHyphen.Pattern = "--"

    Dim enDash As String
    enDash = ChrW(8211)
    Dim emDash As String
    emDash = ChrW(8212)

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDash

        If Not EngineIsInPageRange(paraRange) Then GoTo NextParaDash

        paraText = paraRange.Text
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDash
        ' Strip para mark
        If Len(paraText) > 0 Then
            If Right$(paraText, 1) = vbCr Or Right$(paraText, 1) = Chr(13) Then
                paraText = Left$(paraText, Len(paraText) - 1)
            End If
        End If
        If Len(paraText) < 2 Then GoTo NextParaDash

        ' Calculate auto-number prefix offset
        Dim dashListPrefixLen As Long
        dashListPrefixLen = GetDashListPrefixLen(para, paraText)

        ' --- Check 1: Hyphen in number ranges (digit-digit) ---
        Dim mHR As Object
        Set mHR = reHyphenRange.Execute(paraText)
        Dim hm As Object
        For Each hm In mHR
            Dim hrStart As Long
            hrStart = paraRange.Start + hm.FirstIndex - dashListPrefixLen
            ' The hyphen is at offset +length_of_first_digit
            ' In pattern (\d)-(\d), hyphen is at FirstIndex + 1
            Dim hyphenPos As Long
            hyphenPos = hrStart + 1
            Dim hrEnd As Long
            hrEnd = hyphenPos + 1  ' just the hyphen

            Err.Clear
            Dim hrRng As Range
            Set hrRng = doc.Range(hyphenPos, hrEnd)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            Else
                locStr = EngineGetLocationString(hrRng, doc)
            End If

            Set finding = CreateIssueDict(RULE_NAME_DASH, locStr, _
                "Hyphen used in number range. Use an en-dash (" & enDash & ") for ranges.", _
                enDash, hyphenPos, hrEnd, "error", True)
            issues.Add finding
        Next hm

        ' --- Check 2: Double-hyphen "--" should be em-dash ---
        Dim mDH As Object
        Set mDH = reDoubleHyphen.Execute(paraText)
        Dim dhm As Object
        For Each dhm In mDH
            Dim dhStart As Long
            dhStart = paraRange.Start + dhm.FirstIndex - dashListPrefixLen
            Dim dhEnd As Long
            dhEnd = dhStart + 2

            Err.Clear
            Dim dhRng As Range
            Set dhRng = doc.Range(dhStart, dhEnd)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            Else
                locStr = EngineGetLocationString(dhRng, doc)
            End If

            Set finding = CreateIssueDict(RULE_NAME_DASH, locStr, _
                "Double-hyphen found. Use an em-dash (" & emDash & ") instead.", _
                emDash, dhStart, dhEnd, "error", True)
            issues.Add finding
        Next dhm

        ' --- Check 3: En-dash between letters (compound word) ---
        ' Pattern: letter + en-dash + letter (no spaces) = should be hyphen
        Dim enPos As Long
        enPos = InStr(1, paraText, enDash)
        Do While enPos > 0
            If enPos > 1 And enPos < Len(paraText) Then
                Dim chBefore As String
                Dim chAfter As String
                chBefore = Mid$(paraText, enPos - 1, 1)
                chAfter = Mid$(paraText, enPos + 1, 1)

                Dim beforeIsLetter As Boolean
                Dim afterIsLetter As Boolean
                beforeIsLetter = (chBefore >= "A" And chBefore <= "Z") Or _
                                 (chBefore >= "a" And chBefore <= "z")
                afterIsLetter = (chAfter >= "A" And chAfter <= "Z") Or _
                                (chAfter >= "a" And chAfter <= "z")

                If beforeIsLetter And afterIsLetter Then
                    Dim enStart As Long
                    enStart = paraRange.Start + enPos - 1 - dashListPrefixLen
                    Dim enEnd As Long
                    enEnd = enStart + 1

                    Err.Clear
                    Dim enRng As Range
                    Set enRng = doc.Range(enStart, enEnd)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    Else
                        locStr = EngineGetLocationString(enRng, doc)
                    End If

                    Set finding = CreateIssueDict(RULE_NAME_DASH, locStr, _
                        "En-dash (" & enDash & ") used between words. Use a hyphen (-) for compound words.", _
                        "-", enStart, enEnd, "error", True)
                    issues.Add finding
                End If

                ' Check 4: Spaced en-dash (" – ") -> should be em-dash (" — ")
                ' Exception: spaced en-dash between numbers is correct for ranges
                Dim beforeIsSpace As Boolean
                Dim afterIsSpace As Boolean
                beforeIsSpace = (chBefore = " ")
                afterIsSpace = (chAfter = " ")

                If beforeIsSpace And afterIsSpace Then
                    ' Check if this is a number range (digit before space and digit after space)
                    Dim isNumberRange As Boolean
                    isNumberRange = False
                    If enPos > 2 And enPos + 1 < Len(paraText) Then
                        Dim charBeforeSpace As String
                        Dim charAfterSpace As String
                        charBeforeSpace = Mid$(paraText, enPos - 2, 1)
                        charAfterSpace = Mid$(paraText, enPos + 2, 1)
                        If (charBeforeSpace >= "0" And charBeforeSpace <= "9") And _
                           (charAfterSpace >= "0" And charAfterSpace <= "9") Then
                            isNumberRange = True
                        End If
                    End If
                    If isNumberRange Then GoTo NextEnDashPos
                    Dim snStart As Long
                    snStart = paraRange.Start + enPos - 1 - dashListPrefixLen
                    Dim snEnd As Long
                    snEnd = snStart + 1

                    Err.Clear
                    Dim snRng As Range
                    Set snRng = doc.Range(snStart, snEnd)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    Else
                        locStr = EngineGetLocationString(snRng, doc)
                    End If

                    Set finding = CreateIssueDict(RULE_NAME_DASH, locStr, _
                        "Spaced en-dash (" & enDash & ") found. Consider using an em-dash (" & emDash & ") for parenthetical interruptions.", _
                        emDash, snStart, snEnd, "warning", False)
                    issues.Add finding
                End If
            End If

NextEnDashPos:
            enPos = InStr(enPos + 1, paraText, enDash)
        Loop

NextParaDash:
    Next para
    On Error GoTo 0

    Set Check_DashUsage = issues
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_Quotes.bas

```vb
Attribute VB_Name = "Rules_Quotes"
' ============================================================
' Rules_Quotes.bas
' Quotation-mark rules for UK legal proofreading:
'   Rule 17: quotation mark consistency (straight vs smart)
'   Rule 32: single quotes as default outer marks (nesting-aware)
'   Rule 33: smart quote consistency (prefers smart)
'
' Performance notes:
'   - All character scanning uses byte arrays, not Mid$/AscW
'   - Apostrophe detection is inlined on the byte data
'   - Rule 17 collects positions in one pass, flags from the array
'   - Rule 33 merged from two paragraph passes into one
'   - Location ranges are reused via SetRange, not re-created
'
' Dependency: PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

' -- Rule name constants ----------------------------------------
Private Const RULE17 As String = "quotation_mark_consistency"
Private Const RULE32 As String = "single_quotes_default"
Private Const RULE33 As String = "smart_quote_consistency"

' -- Unicode code points ----------------------------------------
Private Const QD  As Long = 34     ' straight double  "
Private Const QDO As Long = 8220   ' smart double open
Private Const QDC As Long = 8221   ' smart double close
Private Const QS  As Long = 39     ' straight single  '
Private Const QSO As Long = 8216   ' smart single open
Private Const QSC As Long = 8217   ' smart single close

' ================================================================
'  RULE 17 -- QUOTATION MARK CONSISTENCY
'
'  One byte-array pass over doc.Content.Text to count + collect
'  positions of every quote type.  Determines dominant style for
'  doubles and singles independently (ties -> straight).  Emits
'  findings for each minority occurrence within the page range.
' ================================================================
Public Function Check_QuotationMarkConsistency( _
        doc As Document) As Collection

    Dim issues As New Collection

    ' -- Grab full-document text once ---------------------------
    Dim docText As String
    On Error Resume Next
    docText = doc.Content.Text
    If Err.Number <> 0 Then
        Err.Clear: On Error GoTo 0
        Set Check_QuotationMarkConsistency = issues
        Exit Function
    End If
    On Error GoTo 0

    If LenB(docText) = 0 Then
        Set Check_QuotationMarkConsistency = issues
        Exit Function
    End If

    ' -- Convert to byte array for fast scanning ----------------
    '    VBA strings are UTF-16LE: two bytes per character.
    '    Byte(i) is low byte, Byte(i+1) is high byte.
    '    Character's document position = i \ 2.
    Dim b() As Byte
    b = docText
    Dim bMax As Long
    bMax = UBound(b) - 1   ' last even index

    ' -- Counters -----------------------------------------------
    Dim cSD As Long   ' straight double
    Dim cCD As Long   ' smart double
    Dim cSS As Long   ' straight single (excluding apostrophes)
    Dim cCS As Long   ' smart single   (excluding apostrophes)

    ' -- Position collectors (grow-on-demand) --------------------
    Dim pSD() As Long
    ReDim pSD(0 To 127)
    Dim pCD() As Long
    ReDim pCD(0 To 127)
    Dim pSS() As Long
    ReDim pSS(0 To 127)
    Dim pCS() As Long
    ReDim pCS(0 To 127)
    Dim capSD As Long
    capSD = 128
    Dim capCD As Long
    capCD = 128
    Dim capSS As Long
    capSS = 128
    Dim capCS As Long
    capCS = 128

    ' -- Single pass: count + collect positions ------------------
    Dim i As Long, code As Long

    For i = 0 To bMax Step 2
        code = b(i) Or (CLng(b(i + 1)) * 256&)

        Select Case code
        Case QD
            If cSD >= capSD Then
                capSD = capSD * 2
                ReDim Preserve pSD(0 To capSD - 1)
            End If
            pSD(cSD) = i \ 2: cSD = cSD + 1

        Case QDO, QDC
            If cCD >= capCD Then
                capCD = capCD * 2
                ReDim Preserve pCD(0 To capCD - 1)
            End If
            pCD(cCD) = i \ 2: cCD = cCD + 1

        Case QS
            If Not ByteIsApostrophe(b, i, bMax) Then
                If cSS >= capSS Then
                    capSS = capSS * 2
                    ReDim Preserve pSS(0 To capSS - 1)
                End If
                pSS(cSS) = i \ 2: cSS = cSS + 1
            End If

        Case QSO
            If cCS >= capCS Then
                capCS = capCS * 2
                ReDim Preserve pCS(0 To capCS - 1)
            End If
            pCS(cCS) = i \ 2: cCS = cCS + 1

        Case QSC
            If Not ByteIsApostrophe(b, i, bMax) Then
                If cCS >= capCS Then
                    capCS = capCS * 2
                    ReDim Preserve pCS(0 To capCS - 1)
                End If
                pCS(cCS) = i \ 2: cCS = cCS + 1
            End If
        End Select
    Next i

    ' -- Determine dominant styles (tie -> straight) ------------
    Dim dblStraight As Boolean
    dblStraight = (cSD >= cCD)
    Dim sglStraight As Boolean
    sglStraight = (cSS >= cCS)

    ' -- Flag minority doubles ----------------------------------
    If dblStraight And cCD > 0 Then
        EmitFromPositions doc, issues, pCD, cCD, RULE17, _
            "Smart double quotation mark found; " & _
            "document predominantly uses straight", _
            "Change to straight double quotation mark (" & _
            Chr$(QD) & ")"

    ElseIf (Not dblStraight) And cSD > 0 Then
        EmitFromPositions doc, issues, pSD, cSD, RULE17, _
            "Straight double quotation mark found; " & _
            "document predominantly uses smart", _
            "Change to smart double quotation marks (" & _
            ChrW$(QDO) & ChrW$(QDC) & ")"
    End If

    ' -- Flag minority singles ----------------------------------
    If sglStraight And cCS > 0 Then
        EmitFromPositions doc, issues, pCS, cCS, RULE17, _
            "Smart single quotation mark found; " & _
            "document predominantly uses straight", _
            "Change to straight single quotation mark (" & _
            Chr$(QS) & ")"

    ElseIf (Not sglStraight) And cSS > 0 Then
        EmitFromPositions doc, issues, pSS, cSS, RULE17, _
            "Straight single quotation mark found; " & _
            "document predominantly uses smart", _
            "Change to smart single quotation marks (" & _
            ChrW$(QSO) & ChrW$(QSC) & ")"
    End If

    Set Check_QuotationMarkConsistency = issues
End Function

' ================================================================
'  RULE 32 -- SINGLE / DOUBLE QUOTES DEFAULT (nesting-aware)
'
'  Proper nesting-aware scanner that classifies each non-apostrophe
'  quote character as outer-level or inner-level, then only flags
'  quotes that use the wrong type AT THEIR NESTING LEVEL.
'
'  UK convention (nestMode="SINGLE"):
'    Depth 0 (outer) should use single quotes.
'    Depth 1 (inner) should use double quotes.
'    Any deeper level alternates.
'
'  The scanner uses explicit nesting stacks per paragraph.
'  Apostrophes (letter-flanked ' chars) are always skipped.
'  Straight quotes toggle; smart quotes have open/close directionality.
' ================================================================
Public Function Check_SingleQuotesDefault( _
        doc As Document) As Collection

    Dim issues As New Collection
    Dim para As Paragraph
    Dim pRng As Range
    Dim pText As String
    Dim pStart As Long
    Dim styleName As String
    Dim b() As Byte
    Dim bMax As Long
    Dim i As Long, code As Long, pos As Long
    Dim locStr As String

    ' Determine which quote type to flag based on user preference
    Dim nestMode As String
    nestMode = EngineGetQuoteNesting()  ' "SINGLE" or "DOUBLE"

    ' Outer and inner quote codes based on nesting mode
    ' For SINGLE outer: even depths (0,2,4..) use single, odd depths use double
    ' For DOUBLE outer: even depths use double, odd depths use single
    Dim outerOpen As Long, outerClose As Long, outerStraight As Long
    Dim innerOpen As Long, innerClose As Long, innerStraight As Long
    If nestMode = "DOUBLE" Then
        outerOpen = QDO: outerClose = QDC: outerStraight = QD
        innerOpen = QSO: innerClose = QSC: innerStraight = QS
    Else
        outerOpen = QSO: outerClose = QSC: outerStraight = QS
        innerOpen = QDO: innerClose = QDC: innerStraight = QD
    End If

    ' Reusable range -- created once, repositioned via SetRange
    Dim locRng As Range
    Set locRng = doc.Range(0, 1)

    ' Nesting state persists across paragraphs for multi-paragraph quotes
    Dim nestDepth As Long
    nestDepth = 0
    ' Track what type each nesting level was opened with:
    ' True = outer-type, False = inner-type.  Max 10 levels.
    Dim levelIsOuter(0 To 9) As Boolean
    ' Count of nesting anomalies (underflow / overflow).
    ' If too many anomalies we suppress further flagging.
    Dim nestAnomalies As Long
    nestAnomalies = 0

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set pRng = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NxtP32

        ' Page-range gate (once per paragraph, not per character)
        If Not EngineIsInPageRange(pRng) Then GoTo NxtP32

        ' Style exclusion gate
        Err.Clear
        styleName = pRng.ParagraphStyle
        If Err.Number <> 0 Then styleName = "": Err.Clear
        If IsExcludedStyle(styleName) Then GoTo NxtP32

        ' Fetch paragraph text
        Err.Clear
        pText = pRng.Text
        If Err.Number <> 0 Then Err.Clear: GoTo NxtP32
        If LenB(pText) = 0 Then GoTo NxtP32

        pStart = pRng.Start
        b = pText
        bMax = UBound(b) - 1

        ' Compute list prefix length for position correction
        Dim r32ListPrefixLen As Long
        r32ListPrefixLen = GetQListPrefixLen(para, pText)

        ' ======================================================
        '  NESTING-AWARE SCAN
        '
        '  nestDepth tracks how many quote levels deep we are.
        '  Even depths (0, 2, ...) expect the "outer" type.
        '  Odd  depths (1, 3, ...) expect the "inner" type.
        '
        '  Only FLAG wrong-type OPENING quotes.  Closing quotes
        '  just pop the stack. Straight-quote open/close is
        '  determined by preceding-character context (space/SOL
        '  = opening, letter/digit = closing).
        '
        '  Soft failure: if nestDepth underflows on close, reset
        '  to 0 and increment anomaly counter; if anomaly count
        '  is high, stop flagging for this document.
        ' ======================================================

        For i = 0 To bMax Step 2
            code = b(i) Or (CLng(b(i + 1)) * 256&)

            ' -- Apostrophe skip for any single-quote character --
            Dim isApostrophe As Boolean
            isApostrophe = False
            If code = QS Or code = QSC Or code = QSO Then
                isApostrophe = ByteIsApostropheExt(b, i, bMax)
            End If
            If isApostrophe Then GoTo NxtChar32

            ' -- Classify this character --
            Dim isOuterOpen As Boolean
            Dim isOuterClose As Boolean
            Dim isInnerOpen As Boolean
            Dim isInnerClose As Boolean
            Dim isStraightOuter As Boolean
            Dim isStraightInner As Boolean
            isOuterOpen = (code = outerOpen)
            isOuterClose = (code = outerClose)
            isInnerOpen = (code = innerOpen)
            isInnerClose = (code = innerClose)
            isStraightOuter = (code = outerStraight)
            isStraightInner = (code = innerStraight)

            ' Skip non-quote characters
            If Not (isOuterOpen Or isOuterClose Or isInnerOpen Or _
                    isInnerClose Or isStraightOuter Or isStraightInner) Then
                GoTo NxtChar32
            End If

            ' -- Determine if this is an opening or closing quote --
            Dim isOpening As Boolean
            Dim isClosing As Boolean
            Dim isOuterType As Boolean  ' True = outer-style char

            isOuterType = (isOuterOpen Or isOuterClose Or isStraightOuter)

            ' Smart quotes have directionality built in
            If isOuterOpen Or isInnerOpen Then
                isOpening = True
                isClosing = False
            ElseIf isOuterClose Or isInnerClose Then
                isOpening = False
                isClosing = True
            Else
                ' Straight quote: use preceding-character context
                ' Space, tab, newline, SOL, open-paren = opening
                ' Letter, digit, punctuation = closing
                Dim prevIsSpace As Boolean
                prevIsSpace = True  ' default: start-of-line = opening
                If i >= 2 Then
                    Dim prevCode As Long
                    prevCode = b(i - 2) Or (CLng(b(i - 1)) * 256&)
                    prevIsSpace = (prevCode = 32 Or prevCode = 9 Or _
                                   prevCode = 13 Or prevCode = 10 Or _
                                   prevCode = 160 Or prevCode = 40 Or _
                                   prevCode = 91 Or prevCode = 8212 Or _
                                   prevCode = 8211)
                End If
                If prevIsSpace Then
                    isOpening = True
                    isClosing = False
                Else
                    ' Check if the current nesting level was opened with same type
                    If nestDepth > 0 Then
                        Dim curLevelOuter As Boolean
                        If nestDepth <= 10 Then
                            curLevelOuter = levelIsOuter(nestDepth - 1)
                        Else
                            curLevelOuter = ((nestDepth Mod 2) = 0)
                        End If
                        If curLevelOuter = isOuterType Then
                            isOpening = False
                            isClosing = True
                        Else
                            isOpening = True
                            isClosing = False
                        End If
                    Else
                        ' At depth 0, non-space-preceded straight quote
                        ' is probably closing an untracked opening
                        isOpening = False
                        isClosing = True
                    End If
                End If
            End If

            ' -- Process opening quotes --
            If isOpening Then
                Dim expectOuter As Boolean
                expectOuter = ((nestDepth Mod 2) = 0)

                Dim wrongTypeOpen As Boolean
                wrongTypeOpen = (isOuterType <> expectOuter)

                ' Only flag if we have low anomaly count (reliable nesting state)
                If wrongTypeOpen And nestAnomalies < 5 Then
                    Dim issMsg As String
                    If expectOuter Then
                        If nestMode = "DOUBLE" Then
                            issMsg = "Outer quotation marks should use double quotation marks, not single."
                        Else
                            issMsg = "Outer quotation marks should use single quotation marks, not double."
                        End If
                    Else
                        If nestMode = "DOUBLE" Then
                            issMsg = "Inner quotation marks should use single quotation marks, not double."
                        Else
                            issMsg = "Inner quotation marks should use double quotation marks, not single."
                        End If
                    End If

                    pos = pStart + (i \ 2) - r32ListPrefixLen
                    Err.Clear
                    locRng.SetRange pos, pos + 1
                    If Err.Number <> 0 Then
                        locStr = "unknown location": Err.Clear
                    Else
                        Err.Clear
                        locStr = EngineGetLocationString(locRng, doc)
                        If Err.Number <> 0 Then
                            locStr = "unknown location": Err.Clear
                        End If
                    End If

                    issues.Add CreateIssueDict(RULE32, locStr, _
                        issMsg, "", pos, pos + 1, "warning")
                End If

                ' Push nesting level regardless (so its close is tracked)
                If nestDepth < 10 Then
                    levelIsOuter(nestDepth) = isOuterType
                End If
                nestDepth = nestDepth + 1
                GoTo NxtChar32
            End If

            ' -- Process closing quotes --
            If isClosing Then
                If nestDepth > 0 Then
                    nestDepth = nestDepth - 1
                Else
                    ' Underflow: malformed sequence.  Soft-reset.
                    nestAnomalies = nestAnomalies + 1
                    nestDepth = 0
                End If
                GoTo NxtChar32
            End If

NxtChar32:
        Next i

NxtP32:
    Next para
    On Error GoTo 0

    Set Check_SingleQuotesDefault = issues
End Function

' ================================================================
'  RULE 33 -- SMART QUOTE CONSISTENCY
'
'  Single pass over paragraphs: counts straight vs smart quotes
'  AND collects minority-style positions simultaneously.
'  Preference (smart or straight) is read from the engine toggle.
'  If both styles exist, flags the non-preferred style.
' ================================================================
Public Function Check_SmartQuoteConsistency( _
        doc As Document) As Collection

    Dim issues As New Collection
    Dim para As Paragraph
    Dim pRng As Range
    Dim pText As String
    Dim pStart As Long
    Dim b() As Byte
    Dim bMax As Long
    Dim i As Long, code As Long

    Dim prefStyle As String
    prefStyle = EngineGetSmartQuotePref()  ' "SMART" or "STRAIGHT"
    Dim preferSmart As Boolean
    preferSmart = (prefStyle <> "STRAIGHT")

    ' Counters
    Dim cStraight As Long, cSmart As Long

    ' Collect positions of the non-preferred style
    Dim fPos() As Long
    Dim fCnt As Long, fCap As Long
    fCap = 256
    ReDim fPos(0 To fCap - 1)

    ' -- Single pass: count + collect positions ------------------
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set pRng = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NxtP33

        If Not EngineIsInPageRange(pRng) Then GoTo NxtP33

        Err.Clear
        pText = pRng.Text
        If Err.Number <> 0 Then Err.Clear: GoTo NxtP33
        If LenB(pText) = 0 Then GoTo NxtP33

        pStart = pRng.Start
        b = pText
        bMax = UBound(b) - 1

        ' Compute list prefix length for position correction
        Dim r33ListPrefixLen As Long
        r33ListPrefixLen = GetQListPrefixLen(para, pText)

        For i = 0 To bMax Step 2
            code = b(i) Or (CLng(b(i + 1)) * 256&)

            Select Case code
            Case QD
                cStraight = cStraight + 1
                If Not preferSmart Then GoTo NxtCode33
                ' Prefer smart -> collect straight positions
                If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                fPos(fCnt) = pStart + (i \ 2) - r33ListPrefixLen: fCnt = fCnt + 1

            Case QDO, QDC
                cSmart = cSmart + 1
                If preferSmart Then GoTo NxtCode33
                ' Prefer straight -> collect smart positions
                If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                fPos(fCnt) = pStart + (i \ 2) - r33ListPrefixLen: fCnt = fCnt + 1

            Case QS
                If Not ByteIsApostrophe(b, i, bMax) Then
                    cStraight = cStraight + 1
                    If preferSmart Then
                        If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                        fPos(fCnt) = pStart + (i \ 2) - r33ListPrefixLen: fCnt = fCnt + 1
                    End If
                End If

            Case QSO
                cSmart = cSmart + 1
                If Not preferSmart Then
                    If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                    fPos(fCnt) = pStart + (i \ 2) - r33ListPrefixLen: fCnt = fCnt + 1
                End If

            Case QSC
                If Not ByteIsApostrophe(b, i, bMax) Then
                    cSmart = cSmart + 1
                    If Not preferSmart Then
                        If fCnt >= fCap Then fCap = fCap * 2: ReDim Preserve fPos(0 To fCap - 1)
                        fPos(fCnt) = pStart + (i \ 2) - r33ListPrefixLen: fCnt = fCnt + 1
                    End If
                End If
            End Select
NxtCode33:
        Next i

NxtP33:
    Next para
    On Error GoTo 0

    ' -- No mix? Nothing to report ------------------------------
    If cStraight = 0 Or cSmart = 0 Then
        Set Check_SmartQuoteConsistency = issues
        Exit Function
    End If

    ' -- Summary finding ----------------------------------------
    Dim prefName As String, wrongName As String
    If preferSmart Then prefName = "smart": wrongName = "straight" _
    Else prefName = "straight": wrongName = "smart"

    issues.Add CreateIssueDict(RULE33, "Document", _
        "Quotation mark style is inconsistent. Found " & _
        cStraight & " straight and " & cSmart & _
        " smart quotation marks.", _
        "Use " & prefName & " quotation marks consistently " & _
        "throughout the document.", 0, 0, "warning")

    ' -- Flag each non-preferred quote ---------------------------
    Dim locRng As Range
    Set locRng = doc.Range(0, 1)

    Dim j As Long, pos As Long, locStr As String
    On Error Resume Next
    For j = 0 To fCnt - 1
        pos = fPos(j)
        Err.Clear
        locRng.SetRange pos, pos + 1
        If Err.Number <> 0 Then Err.Clear: GoTo SkipP33

        Err.Clear
        locStr = EngineGetLocationString(locRng, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear

        issues.Add CreateIssueDict(RULE33, locStr, _
            UCase(Left(wrongName, 1)) & Mid(wrongName, 2) & _
            " quotation mark found in document.", _
            "Replace with " & prefName & " quotation mark.", _
            pos, pos + 1, "warning")
SkipP33:
    Next j
    On Error GoTo 0

    Set Check_SmartQuoteConsistency = issues
End Function

' ================================================================
'  PRIVATE HELPERS
' ================================================================

' ------------------------------------------------------------
'  Apostrophe check on raw byte data (original strict version).
'  True when the character at byte offset bi is flanked by
'  letters on both sides (= mid-word = apostrophe, not quote).
'  Works directly on the byte array -- no Mid$/AscW overhead.
' ------------------------------------------------------------
Private Function ByteIsApostrophe(b() As Byte, _
        ByVal bi As Long, ByVal bMax As Long) As Boolean
    Dim pc As Long, nc As Long
    If bi < 2 Or bi + 3 > bMax Then Exit Function  ' False
    pc = b(bi - 2) Or (CLng(b(bi - 1)) * 256&)
    nc = b(bi + 2) Or (CLng(b(bi + 3)) * 256&)
    ByteIsApostrophe = IsLetterCode(pc) And IsLetterCode(nc)
End Function

' ------------------------------------------------------------
'  Extended apostrophe check: letter or digit flanked.
'  Treats letter+digit and digit+letter combos as apostrophes
'  too (e.g. 90's, '80s).  Used by the nesting scanner.
' ------------------------------------------------------------
Private Function ByteIsApostropheExt(b() As Byte, _
        ByVal bi As Long, ByVal bMax As Long) As Boolean
    Dim pc As Long, nc As Long
    If bi < 2 Or bi + 3 > bMax Then Exit Function  ' False
    pc = b(bi - 2) Or (CLng(b(bi - 1)) * 256&)
    nc = b(bi + 2) Or (CLng(b(bi + 3)) * 256&)
    ' Both sides must be letter or digit
    ByteIsApostropheExt = (IsLetterCode(pc) Or IsDigitCode(pc)) And _
                          (IsLetterCode(nc) Or IsDigitCode(nc))
End Function

' ------------------------------------------------------------
'  Letter test by code point (A-Z, a-z, extended Latin U+00C0
'  through U+02AF).  Covers accented characters common in UK
'  legal text (cafe, naive, resume, etc.).
' ------------------------------------------------------------
Private Function IsLetterCode(ByVal c As Long) As Boolean
    IsLetterCode = (c >= 65 And c <= 90) Or _
                   (c >= 97 And c <= 122) Or _
                   (c >= 192 And c <= 687)
End Function

' ------------------------------------------------------------
'  Digit test by code point (0-9).
' ------------------------------------------------------------
Private Function IsDigitCode(ByVal c As Long) As Boolean
    IsDigitCode = (c >= 48 And c <= 57)
End Function

' ------------------------------------------------------------
'  Style exclusion for Rule 32.  Paragraphs with "Block",
'  "Quote", or "Code" in their style name are skipped.
' ------------------------------------------------------------
Private Function IsExcludedStyle(ByVal sn As String) As Boolean
    If Len(sn) = 0 Then Exit Function  ' False
    Dim ls As String
    ls = LCase$(sn)
    IsExcludedStyle = (InStr(1, ls, "block", vbBinaryCompare) > 0) _
                   Or (InStr(1, ls, "quote", vbBinaryCompare) > 0) _
                   Or (InStr(1, ls, "code", vbBinaryCompare) > 0)
End Function

' ------------------------------------------------------------
'  Emit findings from a pre-collected position array (Rule 17).
'  Uses a single reusable Range to avoid per-finding allocation.
'  Checks page range per position (Rule 17 counts document-wide
'  but only flags within the configured page range).
' ------------------------------------------------------------
Private Sub EmitFromPositions(doc As Document, _
        issues As Collection, _
        positions() As Long, _
        cnt As Long, _
        ruleName As String, _
        issueText As String, _
        suggestion As String)

    If cnt = 0 Then Exit Sub

    Dim locRng As Range
    Set locRng = doc.Range(0, 1)

    Dim j As Long, pos As Long, locStr As String
    On Error Resume Next
    For j = 0 To cnt - 1
        pos = positions(j)
        Err.Clear
        locRng.SetRange pos, pos + 1
        If Err.Number <> 0 Then Err.Clear: GoTo SkipEmit

        If Not EngineIsInPageRange(locRng) Then GoTo SkipEmit

        Err.Clear
        locStr = EngineGetLocationString(locRng, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear

        issues.Add CreateIssueDict(ruleName, locStr, _
            issueText, suggestion, pos, pos + 1, "possible_error")
SkipEmit:
    Next j
    On Error GoTo 0
End Sub

' ------------------------------------------------------------
'  Create a dictionary-based finding (no class dependency).
' ------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
        ByVal location_ As String, _
        ByVal issue_ As String, _
        ByVal suggestion_ As String, _
        ByVal rangeStart_ As Long, _
        ByVal rangeEnd_ As Long, _
        Optional ByVal severity_ As String = "error", _
        Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function

' ------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run( _
        "PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, _
        doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run( _
        "PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ------------------------------------------------------------
'  List prefix length for byte-array position correction.
'  para.Range.Text includes auto-generated list numbering
'  (e.g. "1." & vbTab) but para.Range.Start does NOT account
'  for it, so byte-array positions must subtract this offset.
' ------------------------------------------------------------
Private Function GetQListPrefixLen(para As Paragraph, ByVal paraText As String) As Long
    GetQListPrefixLen = 0
    On Error Resume Next
    Dim lStr As String
    lStr = para.Range.ListFormat.ListString
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    On Error GoTo 0
    If Len(lStr) = 0 Then Exit Function
    ' Verify the text actually starts with the list string
    If Len(paraText) > Len(lStr) Then
        If Left$(paraText, Len(lStr)) = lStr Then
            GetQListPrefixLen = Len(lStr)
            ' Account for tab separator after list number
            If Mid$(paraText, GetQListPrefixLen + 1, 1) = vbTab Then
                GetQListPrefixLen = GetQListPrefixLen + 1
            End If
        End If
    End If
End Function

' ------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetQuoteNesting
' ------------------------------------------------------------
Private Function EngineGetQuoteNesting() As String
    On Error Resume Next
    EngineGetQuoteNesting = Application.Run( _
        "PleadingsEngine.GetQuoteNesting")
    If Err.Number <> 0 Then
        Debug.Print "EngineGetQuoteNesting: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetQuoteNesting = "SINGLE"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetSmartQuotePref
' ------------------------------------------------------------
Private Function EngineGetSmartQuotePref() As String
    On Error Resume Next
    EngineGetSmartQuotePref = Application.Run( _
        "PleadingsEngine.GetSmartQuotePref")
    If Err.Number <> 0 Then
        Debug.Print "EngineGetSmartQuotePref: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetSmartQuotePref = "SMART"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_Spacing.bas

```vb
Attribute VB_Name = "Rules_Spacing"
' ============================================================
' Rules_Spacing.bas
' Spacing and whitespace proofreading rules:
'   - Check_DoubleSpaces      : Flag runs of 2+ spaces
'                                (mode-aware: ONE space or TWO after full stop)
'   - Check_DoubleCommas      : Flag ",," sequences
'   - Check_SpaceBeforePunct  : Flag "word ," patterns
'   - Check_MissingSpaceAfterDot : Flag ".X" (missing space)
'   - Check_TrailingSpaces    : Flag trailing spaces before paragraph marks
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString,
'                          GetSpaceStylePref)
' ============================================================
Option Explicit

Private Const RULE_DOUBLE_SPACES As String = "double_spaces"
Private Const RULE_DOUBLE_COMMAS As String = "double_commas"
Private Const RULE_SPACE_BEFORE_PUNCT As String = "space_before_punct"
Private Const RULE_MISSING_SPACE_DOT As String = "missing_space_after_dot"
Private Const RULE_TRAILING_SPACES As String = "trailing_spaces"

' Known abbreviations (delimited for InStr lookup)
Private Const ABBREV_LIST As String = _
    "|mr|mrs|ms|dr|prof|sr|jr|st|no|nos|" & _
    "|vs|etc|al|approx|dept|govt|inc|ltd|" & _
    "|corp|co|assn|ave|blvd|rd|ct|ft|" & _
    "|vol|rev|gen|sgt|cpl|pvt|lt|capt|" & _
    "|maj|col|cmdr|adm|jan|feb|mar|apr|" & _
    "|jun|jul|aug|sep|oct|nov|dec|mon|" & _
    "|tue|wed|thu|fri|sat|sun|fig|eq|" & _
    "|ref|para|paras|cl|pt|sch|art|reg|v|"

' ============================================================
'  PUBLIC: Check_DoubleSpaces
'  Flags runs of 2+ spaces. In TWO mode, allows double space
'  after sentence-ending full stops (but not after abbreviations).
' ============================================================
Public Function Check_DoubleSpaces(doc As Document) As Collection
    Dim issues As New Collection
    Dim reDouble As Object
    Set reDouble = CreateObject("VBScript.RegExp")
    reDouble.Global = True
    reDouble.Pattern = " {2,}"

    Dim reSingle As Object
    Set reSingle = CreateObject("VBScript.RegExp")
    reSingle.Global = True
    reSingle.Pattern = "\.( )([A-Z])"

    Dim spaceStyle As String
    spaceStyle = EngineGetSpaceStylePref()

    Dim para As Paragraph
    Dim paraText As String
    Dim paraRange As Range
    Dim finding As Object
    Dim locStr As String

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDS

        If Not EngineIsInPageRange(paraRange) Then GoTo NextParaDS

        ' Block quotes are filtered at engine level by FilterBlockQuoteIssues.
        ' Removed per-paragraph Application.Run("IsBlockQuotePara") call here
        ' to eliminate heavy object-model traffic (font/italic/text/style
        ' per paragraph via cross-module dispatch was a major regression cause).

        paraText = StripParaMarkChar(paraRange.Text)
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDS
        If Len(paraText) < 2 Then GoTo NextParaDS

        ' Calculate auto-number prefix offset
        Dim listPrefixLen As Long
        listPrefixLen = GetListPrefixLen(para, paraText)

        ' --- Pass 1: Flag runs of 2+ spaces ---
        Dim mDoubles As Object
        Set mDoubles = reDouble.Execute(paraText)
        Dim md As Object
        For Each md In mDoubles
            Dim mStart As Long
            mStart = md.FirstIndex   ' 0-based

            If spaceStyle = "TWO" And mStart > 0 Then
                ' In two-space mode, double space is correct after sentence-end full stop
                Dim charBefore As String
                charBefore = Mid(paraText, mStart, 1)   ' char at 0-based mStart-1
                If charBefore = "." Then
                    Dim dotPos As Long
                    dotPos = mStart - 1   ' 0-based index of the full stop
                    Dim wb As String
                    wb = GetWordBeforePos(paraText, dotPos)
                    If Not IsLikelyAbbreviation(paraText, dotPos, wb) Then
                        GoTo NextDoubleMatch   ' sentence-end + 2 spaces = correct
                    End If
                End If
            End If

            ' Flag this double space
            Dim dsStart As Long
            Dim dsEnd As Long
            dsStart = paraRange.Start + mStart - listPrefixLen
            dsEnd = dsStart + md.Length

            Err.Clear
            Dim dsRng As Range
            Set dsRng = doc.Range(dsStart, dsEnd)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            Else
                locStr = EngineGetLocationString(dsRng, doc)
            End If

            Dim dsMsg As String
            If md.Length = 2 Then
                dsMsg = "Double space found."
            Else
                dsMsg = md.Length & " consecutive spaces found."
            End If

            ' Range covers only the EXTRA space(s) — keep the first one
            Set finding = CreateIssueDict(RULE_DOUBLE_SPACES, locStr, _
                dsMsg, "", dsStart + 1, dsEnd, "error", True)
            issues.Add finding

NextDoubleMatch:
        Next md

        ' --- Pass 2 (TWO mode only): Flag missing second space after sentence-end ---
        If spaceStyle = "TWO" Then
            Dim mSingles As Object
            Set mSingles = reSingle.Execute(paraText)
            Dim ms As Object
            For Each ms In mSingles
                Dim sdotPos As Long
                sdotPos = ms.FirstIndex   ' 0-based index of the full stop
                Dim swb As String
                swb = GetWordBeforePos(paraText, sdotPos)
                If Not IsLikelyAbbreviation(paraText, sdotPos, swb) Then
                    ' Sentence-end with only one space -- flag it
                    ' Anchor the issue on the full stop + single space
                    Dim msStart As Long
                    msStart = paraRange.Start + sdotPos - listPrefixLen
                    Dim msEnd As Long
                    msEnd = msStart + 2  ' full stop + space

                    Err.Clear
                    Dim msRng As Range
                    Set msRng = doc.Range(msStart, msEnd)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    Else
                        locStr = EngineGetLocationString(msRng, doc)
                    End If

                    ' Suggestion replaces ". " with ".  " (insert extra space)
                    Set finding = CreateIssueDict(RULE_DOUBLE_SPACES, locStr, _
                        "Missing second space after sentence-ending full stop.", _
                        ".  ", msStart, msEnd, _
                        "warning", True)
                    issues.Add finding
                End If
            Next ms
        End If

NextParaDS:
    Next para
    On Error GoTo 0

    Set Check_DoubleSpaces = issues
End Function

' ============================================================
'  PUBLIC: Check_DoubleCommas
'  Flags ",," sequences in paragraph text.
' ============================================================
Public Function Check_DoubleCommas(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraText As String
    Dim paraRange As Range
    Dim finding As Object
    Dim locStr As String
    Dim pos As Long

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDC

        If Not EngineIsInPageRange(paraRange) Then GoTo NextParaDC

        paraText = paraRange.Text
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaDC

        Dim dcListPrefixLen As Long
        dcListPrefixLen = GetListPrefixLen(para, paraText)

        pos = InStr(1, paraText, ",,")
        Do While pos > 0
            Dim dcStart As Long
            dcStart = paraRange.Start + pos - 1 - dcListPrefixLen
            Dim dcEnd As Long
            dcEnd = dcStart + 2

            Err.Clear
            Dim dcRng As Range
            Set dcRng = doc.Range(dcStart, dcEnd)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            Else
                locStr = EngineGetLocationString(dcRng, doc)
            End If

            Set finding = CreateIssueDict(RULE_DOUBLE_COMMAS, locStr, _
                "Double comma found.", ",", _
                dcStart, dcEnd, "error", True)
            issues.Add finding

            pos = InStr(pos + 2, paraText, ",,")
        Loop

NextParaDC:
    Next para
    On Error GoTo 0

    Set Check_DoubleCommas = issues
End Function

' ============================================================
'  PUBLIC: Check_SpaceBeforePunct
'  Flags "word ," / "word ;" / "word :" etc. patterns.
' ============================================================
Public Function Check_SpaceBeforePunct(doc As Document) As Collection
    Dim issues As New Collection
    Dim rng As Range
    Set rng = doc.Content.Duplicate
    Dim finding As Object
    Dim locStr As String

    On Error Resume Next
    With rng.Find
        .ClearFormatting
        .Text = " [,;:!?]"
        .MatchCase = True
        .MatchWildcards = True
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastPos As Long
    lastPos = -1
    Do While rng.Find.Execute
        If rng.Start <= lastPos Then Exit Do   ' stall guard
        lastPos = rng.Start

        If Not EngineIsInPageRange(rng) Then
            rng.Collapse wdCollapseEnd
            GoTo NextSBP
        End If

        Err.Clear
        locStr = EngineGetLocationString(rng, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear

        Dim punctChar As String
        punctChar = Mid(rng.Text, 2, 1)

        ' Range covers only the space (not the punctuation character)
        Set finding = CreateIssueDict(RULE_SPACE_BEFORE_PUNCT, locStr, _
            "Unexpected space before '" & punctChar & "'.", _
            "", rng.Start, rng.Start + 1, "error", True)
        issues.Add finding

        rng.Collapse wdCollapseEnd
NextSBP:
    Loop
    On Error GoTo 0

    Set Check_SpaceBeforePunct = issues
End Function

' ============================================================
'  PUBLIC: Check_MissingSpaceAfterDot
'  Flags ".X" where X is uppercase and the dot is not an
'  abbreviation full stop. Uses per-paragraph regex scanning.
' ============================================================
Public Function Check_MissingSpaceAfterDot(doc As Document) As Collection
    Dim issues As New Collection
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.Pattern = "\.([A-Z])"

    Dim para As Paragraph
    Dim paraText As String
    Dim paraRange As Range
    Dim finding As Object
    Dim locStr As String

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaMSD

        If Not EngineIsInPageRange(paraRange) Then GoTo NextParaMSD

        paraText = StripParaMarkChar(paraRange.Text)
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaMSD
        If Len(paraText) < 2 Then GoTo NextParaMSD

        Dim msdListPrefixLen As Long
        msdListPrefixLen = GetListPrefixLen(para, paraText)

        Dim matches As Object
        Set matches = re.Execute(paraText)
        Dim m As Object
        For Each m In matches
            Dim dotIdx As Long
            dotIdx = m.FirstIndex   ' 0-based position of the full stop
            Dim wordBefore As String
            wordBefore = GetWordBeforePos(paraText, dotIdx)
            If Not IsLikelyAbbreviation(paraText, dotIdx, wordBefore) Then
                Dim msdStart As Long
                msdStart = paraRange.Start + dotIdx - msdListPrefixLen
                Dim msdEnd As Long
                msdEnd = msdStart + 2   ' "." + capital letter

                Err.Clear
                Dim msdRng As Range
                Set msdRng = doc.Range(msdStart, msdEnd)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                Else
                    locStr = EngineGetLocationString(msdRng, doc)
                End If

                Set finding = CreateIssueDict(RULE_MISSING_SPACE_DOT, locStr, _
                    "Missing space after full stop before '" & _
                    Mid(paraText, dotIdx + 2, 1) & "'.", _
                    "Insert a space after the full stop.", _
                    msdStart, msdEnd, "error", False)
                issues.Add finding
            End If
        Next m

NextParaMSD:
    Next para
    On Error GoTo 0

    Set Check_MissingSpaceAfterDot = issues
End Function

' ============================================================
'  PUBLIC: Check_TrailingSpaces
'  Flags trailing spaces before paragraph marks.
' ============================================================
Public Function Check_TrailingSpaces(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraRange As Range
    Dim finding As Object
    Dim locStr As String

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaTS

        If Not EngineIsInPageRange(paraRange) Then GoTo NextParaTS

        Dim t As String
        t = StripParaMarkChar(paraRange.Text)
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaTS

        If Len(t) > 0 And Right$(t, 1) = " " Then
            ' Count trailing spaces
            Dim numSpaces As Long
            numSpaces = 0
            Dim j As Long
            For j = Len(t) To 1 Step -1
                If Mid$(t, j, 1) = " " Then
                    numSpaces = numSpaces + 1
                Else
                    Exit For
                End If
            Next j

            If numSpaces > 0 Then
                ' Trailing spaces sit just before the paragraph mark
                Dim tsStart As Long
                tsStart = paraRange.End - 1 - numSpaces
                Dim tsEnd As Long
                tsEnd = paraRange.End - 1

                Err.Clear
                Dim tsRng As Range
                Set tsRng = doc.Range(tsStart, tsEnd)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                Else
                    locStr = EngineGetLocationString(tsRng, doc)
                End If

                Dim tsMsg As String
                If numSpaces = 1 Then
                    tsMsg = "Trailing space at end of paragraph."
                Else
                    tsMsg = numSpaces & " trailing spaces at end of paragraph."
                End If

                Set finding = CreateIssueDict(RULE_TRAILING_SPACES, locStr, _
                    tsMsg, "", _
                    tsStart, tsEnd, "warning", True)
                issues.Add finding
            End If
        End If

NextParaTS:
    Next para
    On Error GoTo 0

    Set Check_TrailingSpaces = issues
End Function

' ============================================================
'  PRIVATE HELPERS
' ============================================================

' Calculate the length of auto-generated list numbering text
' that appears in Range.Text but doesn't map to document positions.
' Returns 0 for non-list paragraphs.
Private Function GetListPrefixLen(para As Paragraph, ByVal paraText As String) As Long
    GetListPrefixLen = 0
    On Error Resume Next
    Dim lStr As String
    lStr = para.Range.ListFormat.ListString
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    If Len(lStr) = 0 Then On Error GoTo 0: Exit Function
    ' Verify the text actually starts with the list string
    If Len(paraText) > Len(lStr) Then
        If Left$(paraText, Len(lStr)) = lStr Then
            GetListPrefixLen = Len(lStr)
            ' Account for tab separator after list number
            If Mid$(paraText, GetListPrefixLen + 1, 1) = vbTab Then
                GetListPrefixLen = GetListPrefixLen + 1
            End If
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Strip the trailing paragraph mark (vbCr / Chr(13)) from text
Private Function StripParaMarkChar(ByVal txt As String) As String
    If Len(txt) > 0 Then
        If Right$(txt, 1) = vbCr Or Right$(txt, 1) = Chr(13) Then
            txt = Left$(txt, Len(txt) - 1)
        End If
    End If
    StripParaMarkChar = txt
End Function

' Return the word (letters only) immediately before 0-based position pos
Private Function GetWordBeforePos(ByVal s As String, ByVal pos As Long) As String
    Dim result As String
    result = ""
    Dim i As Long
    Dim c As String
    For i = pos - 1 To 0 Step -1
        c = Mid$(s, i + 1, 1)   ' convert 0-based to 1-based for Mid
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Then
            result = c & result
        Else
            Exit For
        End If
    Next i
    GetWordBeforePos = result
End Function

' Check if word is a known abbreviation
Private Function IsAbbrevWord(ByVal word As String) As Boolean
    IsAbbrevWord = (InStr(1, ABBREV_LIST, "|" & LCase(word) & "|", vbTextCompare) > 0)
End Function

' Extended abbreviation detection:
'   1. Known abbreviation list (Mr, Dr, etc, vs ...)
'   2. Dotted abbreviation: wordBefore is 1-2 chars preceded by a dot (e.g. "e" in "i.e.")
'   3. Ellipsis: empty wordBefore preceded by a dot ("...Word")
'   4. First dot of dotted abbreviation: 1-char wordBefore followed by letter+dot
'
' Index arithmetic: pos is 0-based; Mid uses 1-based.
'   char at 0-based N = Mid(s, N+1, 1)
Private Function IsLikelyAbbreviation(ByVal paraText As String, _
                                       ByVal pos As Long, _
                                       ByVal wordBefore As String) As Boolean
    IsLikelyAbbreviation = False

    ' 1. Standard abbreviation list
    If IsAbbrevWord(wordBefore) Then
        IsLikelyAbbreviation = True
        Exit Function
    End If

    ' 1b. Single uppercase letter (initial: "J. Smith", "A. Jones")
    If Len(wordBefore) = 1 Then
        Dim wbCode As Long
        wbCode = AscW(wordBefore)
        If wbCode >= 65 And wbCode <= 90 Then  ' A-Z
            IsLikelyAbbreviation = True
            Exit Function
        End If
    End If

    ' 2. Dotted abbreviation: wordBefore is 1-2 chars and char before it is a dot
    If Len(wordBefore) >= 1 And Len(wordBefore) <= 2 And _
       pos > Len(wordBefore) Then
        If Mid$(paraText, pos - Len(wordBefore), 1) = "." Then
            IsLikelyAbbreviation = True
            Exit Function
        End If
    End If

    ' 3. Ellipsis: empty wordBefore and char before this dot is also a dot
    If Len(wordBefore) = 0 And pos >= 1 Then
        If Mid$(paraText, pos, 1) = "." Then
            IsLikelyAbbreviation = True
            Exit Function
        End If
    End If

    ' 4. First dot of dotted abbreviation: 1-char wordBefore and
    '    char after this dot is letter followed by another dot
    If Len(wordBefore) = 1 And pos + 2 < Len(paraText) Then
        If Mid$(paraText, pos + 2, 1) Like "[A-Za-z]" And _
           Mid$(paraText, pos + 3, 1) = "." Then
            IsLikelyAbbreviation = True
            Exit Function
        End If
    End If
End Function

' ----------------------------------------------------------------
'  Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetSpaceStylePref
' ----------------------------------------------------------------
Private Function EngineGetSpaceStylePref() As String
    On Error Resume Next
    EngineGetSpaceStylePref = Application.Run("PleadingsEngine.GetSpaceStylePref")
    If Err.Number <> 0 Then
        Debug.Print "EngineGetSpaceStylePref: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetSpaceStylePref = "ONE"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_Spelling.bas

```vb
Attribute VB_Name = "Rules_Spelling"
' ============================================================
' Rules_Spelling.bas
' Combined proofreading rules for UK/US English spelling.
'
' Rule 1 -- British/US Spelling:
'   Detects ~133 spelling differences between US and UK English,
'   with a configurable direction (UK or US mode).
'   Categories: -or/-our, -ize/-ise, -ization/-isation,
'   -er/-re, -se/-ce, -og/-ogue, -ment variants, misc.
'
'   Text in italics or inside quotation marks is NOT auto-fixed
'   but is flagged as a "possible_error" for manual review.
'
' Rule 12 -- Licence/License:
'   Checks correct UK usage of licence (noun) vs license (verb).
'   Also handles compounds and derivatives.
'   UK convention:
'     licence = noun ("a licence", "the licence holder")
'     license = verb ("to license", "shall license")
'     licensed, licensing = always -s- (verb derivatives)
'
' Rule 13 -- Colour Formatting:
'   Detects non-standard font colours in the document body.
'   Identifies the dominant text colour and flags any runs
'   using a different colour (excluding hyperlinks and
'   heading-styled paragraphs).
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, IsWhitelistedTerm,
'                          GetLocationString, GetSpellingMode)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "spelling"
Private Const RULE_NAME_LICENCE As String = "licence_license"
Private Const RULE_NAME_COLOUR As String = "colour_formatting"
Private Const RULE_NAME_CHECK As String = "check_cheque"

' ============================================================
'  MAIN ENTRY POINT
' ============================================================
Public Function Check_Spelling(doc As Document) As Collection
    Dim issues As New Collection
    Dim usWords() As String
    Dim ukWords() As String
    Dim searchWords() As String
    Dim targetWords() As String
    Dim exceptions() As String
    Dim spellingMode As String
    Dim direction As String

    ' -- Build the US <-> UK mapping arrays ----------------
    BuildSpellingArrays usWords, ukWords

    ' -- Determine spelling mode -------------------------
    spellingMode = EngineGetSpellingMode()

    If spellingMode = "US" Then
        ' Search for UK words, suggest US replacements
        searchWords = ukWords
        targetWords = usWords
        direction = "US"

        ' In US mode, no special legal exceptions
        exceptions = Split("program,practice", ",")
    Else
        ' Default: "UK" -- search for US words, suggest UK replacements
        searchWords = usWords
        targetWords = ukWords
        direction = "UK"

        ' "judgment" is standard in UK legal writing (not "judgement")
        ' "practice" is the correct UK noun form (verb: "practise")
        exceptions = Split("program,judgment,practice", ",")
    End If

    ' -- Search main document body -----------------------
    SearchRangeForSpellingIssues doc.Content, doc, searchWords, targetWords, exceptions, direction, issues

    ' -- Search footnotes --------------------------------
    On Error Resume Next
    Dim fn As Footnote
    For Each fn In doc.Footnotes
        Err.Clear
        SearchRangeForSpellingIssues fn.Range, doc, searchWords, targetWords, exceptions, direction, issues
        If Err.Number <> 0 Then Err.Clear
    Next fn
    On Error GoTo 0

    ' -- Search endnotes ---------------------------------
    On Error Resume Next
    Dim en As Endnote
    For Each en In doc.Endnotes
        Err.Clear
        SearchRangeForSpellingIssues en.Range, doc, searchWords, targetWords, exceptions, direction, issues
        If Err.Number <> 0 Then Err.Clear
    Next en
    On Error GoTo 0

    Set Check_Spelling = issues
End Function

' ============================================================
'  PRIVATE: Search a Range for spelling issues
'  Iterates every search/target pair, uses Word's Find to
'  locate whole-word, case-insensitive matches, then filters
'  by page range and whitelist before creating issues.
'
'  direction = "UK" or "US" -- controls the finding text:
'    "UK" -> "US spelling detected: '...'"
'    "US" -> "UK spelling detected: '...'"
' ============================================================
Private Sub SearchRangeForSpellingIssues(searchRange As Range, _
                                         doc As Document, _
                                         ByRef searchWords() As String, _
                                         ByRef targetWords() As String, _
                                         ByRef exceptions() As String, _
                                         ByVal direction As String, _
                                         ByRef issues As Collection)
    Dim i As Long
    Dim rng As Range
    Dim foundText As String
    Dim finding As Object
    Dim locStr As String
    Dim issueText As String
    Dim sourceLabel As String

    ' Determine the label for the detected spelling variant
    If direction = "UK" Then
        sourceLabel = "US"
    Else
        sourceLabel = "UK"
    End If

    For i = LBound(searchWords) To UBound(searchWords)

        ' Reset a fresh range for each search term
        On Error Resume Next
        Set rng = searchRange.Duplicate
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextWord
        End If
        On Error GoTo 0

        With rng.Find
            .ClearFormatting
            .Text = searchWords(i)
            .MatchWholeWord = True
            .MatchCase = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
        End With

        ' Loop through all occurrences of this term
        Do
            On Error Resume Next
            Dim found As Boolean
            found = rng.Find.Execute
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit Do
            End If
            On Error GoTo 0

            If Not found Then Exit Do

            foundText = rng.Text

            ' -- Skip exceptions -----------------------
            If IsException(foundText, exceptions) Then
                GoTo ContinueSearch
            End If

            ' -- Skip whitelisted terms ----------------
            If EngineIsWhitelistedTerm(foundText) Then
                GoTo ContinueSearch
            End If

            ' -- Skip if outside configured page range -
            If Not EngineIsInPageRange(rng) Then
                GoTo ContinueSearch
            End If

            ' -- Create the finding ----------------------
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If
            On Error GoTo 0

            issueText = sourceLabel & " spelling detected: '" & foundText & "'"

            ' -- Downgrade italic / quoted text -------
            Dim severity As String
            Dim suggestion As String
            severity = "error"
            suggestion = targetWords(i)

            If IsRangeItalic(rng) Then
                severity = "possible_error"
                suggestion = ""
                issueText = issueText & " (in italic text -- review manually)"
            ElseIf IsInsideQuotes(rng, doc) Then
                severity = "possible_error"
                suggestion = ""
                issueText = issueText & " (in quoted text -- review manually)"
            End If

            Set finding = CreateIssueDict(RULE_NAME, locStr, issueText, suggestion, rng.Start, rng.End, severity)
            issues.Add finding

ContinueSearch:
            ' Collapse range to end of current match to find next
            On Error Resume Next
            rng.Collapse wdCollapseEnd
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                Exit Do
            End If
            On Error GoTo 0
        Loop

NextWord:
    Next i
End Sub

' ============================================================
'  PRIVATE: Check if a found term is in the exceptions list
' ============================================================
Private Function IsException(ByVal term As String, _
                              ByRef exceptions() As String) As Boolean
    Dim i As Long
    Dim lTerm As String
    lTerm = LCase(Trim(term))

    For i = LBound(exceptions) To UBound(exceptions)
        If LCase(Trim(exceptions(i))) = lTerm Then
            IsException = True
            Exit Function
        End If
    Next i

    IsException = False
End Function

' ============================================================
'  PRIVATE: Build the parallel US/UK spelling arrays
'  ~95 pairs across all categories.
' ============================================================
Private Sub BuildSpellingArrays(ByRef usWords() As String, _
                                 ByRef ukWords() As String)
    ' Dynamic pair building -- no hard-coded PAIR_COUNT needed.
    ' Only low-risk, non-contentious US-to-UK variants are included.
    ' Excluded: check/cheque, practice/practise, license/licence,
    '   judgment/judgement, program/programme, draft/draught,
    '   tire/tyre, curb/kerb, story/storey, meter/metre,
    '   sulphur/sulfur, medical/scientific variants.
    Dim pairCount As Long
    pairCount = 0
    ReDim usWords(0 To 255)
    ReDim ukWords(0 To 255)

    ' -- -or -> -our and inflections --
    AddSpellingPair usWords, ukWords, pairCount, "color", "colour"
    AddSpellingPair usWords, ukWords, pairCount, "colors", "colours"
    AddSpellingPair usWords, ukWords, pairCount, "colored", "coloured"
    AddSpellingPair usWords, ukWords, pairCount, "coloring", "colouring"
    AddSpellingPair usWords, ukWords, pairCount, "favor", "favour"
    AddSpellingPair usWords, ukWords, pairCount, "favored", "favoured"
    AddSpellingPair usWords, ukWords, pairCount, "favoring", "favouring"
    AddSpellingPair usWords, ukWords, pairCount, "favorite", "favourite"
    AddSpellingPair usWords, ukWords, pairCount, "favorites", "favourites"
    AddSpellingPair usWords, ukWords, pairCount, "honor", "honour"
    AddSpellingPair usWords, ukWords, pairCount, "honors", "honours"
    AddSpellingPair usWords, ukWords, pairCount, "honored", "honoured"
    AddSpellingPair usWords, ukWords, pairCount, "honoring", "honouring"
    AddSpellingPair usWords, ukWords, pairCount, "humor", "humour"
    AddSpellingPair usWords, ukWords, pairCount, "labor", "labour"
    AddSpellingPair usWords, ukWords, pairCount, "labored", "laboured"
    AddSpellingPair usWords, ukWords, pairCount, "laboring", "labouring"
    AddSpellingPair usWords, ukWords, pairCount, "neighbor", "neighbour"
    AddSpellingPair usWords, ukWords, pairCount, "neighbors", "neighbours"
    AddSpellingPair usWords, ukWords, pairCount, "neighboring", "neighbouring"
    AddSpellingPair usWords, ukWords, pairCount, "neighborhood", "neighbourhood"
    AddSpellingPair usWords, ukWords, pairCount, "behavior", "behaviour"
    AddSpellingPair usWords, ukWords, pairCount, "behaviors", "behaviours"
    AddSpellingPair usWords, ukWords, pairCount, "behavioral", "behavioural"
    AddSpellingPair usWords, ukWords, pairCount, "endeavor", "endeavour"
    AddSpellingPair usWords, ukWords, pairCount, "endeavored", "endeavoured"
    AddSpellingPair usWords, ukWords, pairCount, "endeavoring", "endeavouring"
    AddSpellingPair usWords, ukWords, pairCount, "harbor", "harbour"
    AddSpellingPair usWords, ukWords, pairCount, "harbors", "harbours"
    AddSpellingPair usWords, ukWords, pairCount, "vigor", "vigour"
    AddSpellingPair usWords, ukWords, pairCount, "valor", "valour"
    AddSpellingPair usWords, ukWords, pairCount, "candor", "candour"
    AddSpellingPair usWords, ukWords, pairCount, "clamor", "clamour"
    AddSpellingPair usWords, ukWords, pairCount, "glamor", "glamour"
    AddSpellingPair usWords, ukWords, pairCount, "parlor", "parlour"
    AddSpellingPair usWords, ukWords, pairCount, "rancor", "rancour"
    AddSpellingPair usWords, ukWords, pairCount, "rigor", "rigour"
    AddSpellingPair usWords, ukWords, pairCount, "rumor", "rumour"
    AddSpellingPair usWords, ukWords, pairCount, "rumors", "rumours"
    AddSpellingPair usWords, ukWords, pairCount, "savior", "saviour"
    AddSpellingPair usWords, ukWords, pairCount, "splendor", "splendour"
    AddSpellingPair usWords, ukWords, pairCount, "tumor", "tumour"
    AddSpellingPair usWords, ukWords, pairCount, "tumors", "tumours"
    AddSpellingPair usWords, ukWords, pairCount, "vapor", "vapour"
    AddSpellingPair usWords, ukWords, pairCount, "fervor", "fervour"
    AddSpellingPair usWords, ukWords, pairCount, "armor", "armour"
    AddSpellingPair usWords, ukWords, pairCount, "armored", "armoured"
    AddSpellingPair usWords, ukWords, pairCount, "flavor", "flavour"
    AddSpellingPair usWords, ukWords, pairCount, "flavors", "flavours"
    AddSpellingPair usWords, ukWords, pairCount, "flavored", "flavoured"
    AddSpellingPair usWords, ukWords, pairCount, "flavoring", "flavouring"

    ' -- -er -> -re where generally safe --
    AddSpellingPair usWords, ukWords, pairCount, "center", "centre"
    AddSpellingPair usWords, ukWords, pairCount, "centers", "centres"
    AddSpellingPair usWords, ukWords, pairCount, "centered", "centred"
    AddSpellingPair usWords, ukWords, pairCount, "centering", "centring"
    AddSpellingPair usWords, ukWords, pairCount, "fiber", "fibre"
    AddSpellingPair usWords, ukWords, pairCount, "fibers", "fibres"
    AddSpellingPair usWords, ukWords, pairCount, "theater", "theatre"
    AddSpellingPair usWords, ukWords, pairCount, "theaters", "theatres"
    AddSpellingPair usWords, ukWords, pairCount, "somber", "sombre"
    AddSpellingPair usWords, ukWords, pairCount, "caliber", "calibre"
    AddSpellingPair usWords, ukWords, pairCount, "saber", "sabre"
    AddSpellingPair usWords, ukWords, pairCount, "specter", "spectre"
    AddSpellingPair usWords, ukWords, pairCount, "meager", "meagre"
    AddSpellingPair usWords, ukWords, pairCount, "luster", "lustre"
    AddSpellingPair usWords, ukWords, pairCount, "maneuver", "manoeuvre"
    AddSpellingPair usWords, ukWords, pairCount, "maneuvered", "manoeuvred"
    AddSpellingPair usWords, ukWords, pairCount, "maneuvering", "manoeuvring"
    AddSpellingPair usWords, ukWords, pairCount, "reconnoiter", "reconnoitre"
    AddSpellingPair usWords, ukWords, pairCount, "goiter", "goitre"
    AddSpellingPair usWords, ukWords, pairCount, "ocher", "ochre"

    ' -- -se -> -ce where safe --
    AddSpellingPair usWords, ukWords, pairCount, "defense", "defence"
    AddSpellingPair usWords, ukWords, pairCount, "defenses", "defences"
    AddSpellingPair usWords, ukWords, pairCount, "offense", "offence"
    AddSpellingPair usWords, ukWords, pairCount, "offenses", "offences"
    AddSpellingPair usWords, ukWords, pairCount, "pretense", "pretence"

    ' -- -og -> -ogue --
    AddSpellingPair usWords, ukWords, pairCount, "analog", "analogue"
    AddSpellingPair usWords, ukWords, pairCount, "catalog", "catalogue"
    AddSpellingPair usWords, ukWords, pairCount, "dialog", "dialogue"
    AddSpellingPair usWords, ukWords, pairCount, "monolog", "monologue"
    AddSpellingPair usWords, ukWords, pairCount, "prolog", "prologue"
    AddSpellingPair usWords, ukWords, pairCount, "epilog", "epilogue"

    ' -- -ment and similar safe variants --
    AddSpellingPair usWords, ukWords, pairCount, "acknowledgment", "acknowledgement"
    AddSpellingPair usWords, ukWords, pairCount, "acknowledgments", "acknowledgements"
    AddSpellingPair usWords, ukWords, pairCount, "fulfillment", "fulfilment"
    AddSpellingPair usWords, ukWords, pairCount, "fulfill", "fulfil"
    AddSpellingPair usWords, ukWords, pairCount, "enrollment", "enrolment"
    AddSpellingPair usWords, ukWords, pairCount, "enroll", "enrol"
    AddSpellingPair usWords, ukWords, pairCount, "installment", "instalment"
    AddSpellingPair usWords, ukWords, pairCount, "installments", "instalments"

    ' -- Doubled consonant variants --
    AddSpellingPair usWords, ukWords, pairCount, "traveled", "travelled"
    AddSpellingPair usWords, ukWords, pairCount, "traveling", "travelling"
    AddSpellingPair usWords, ukWords, pairCount, "traveler", "traveller"
    AddSpellingPair usWords, ukWords, pairCount, "travelers", "travellers"
    AddSpellingPair usWords, ukWords, pairCount, "canceled", "cancelled"
    AddSpellingPair usWords, ukWords, pairCount, "canceling", "cancelling"
    AddSpellingPair usWords, ukWords, pairCount, "labeled", "labelled"
    AddSpellingPair usWords, ukWords, pairCount, "labeling", "labelling"
    AddSpellingPair usWords, ukWords, pairCount, "modeled", "modelled"
    AddSpellingPair usWords, ukWords, pairCount, "modeling", "modelling"
    AddSpellingPair usWords, ukWords, pairCount, "counselor", "counsellor"
    AddSpellingPair usWords, ukWords, pairCount, "counselors", "counsellors"
    AddSpellingPair usWords, ukWords, pairCount, "counseling", "counselling"
    AddSpellingPair usWords, ukWords, pairCount, "signaled", "signalled"
    AddSpellingPair usWords, ukWords, pairCount, "signaling", "signalling"
    AddSpellingPair usWords, ukWords, pairCount, "fueled", "fuelled"
    AddSpellingPair usWords, ukWords, pairCount, "fueling", "fuelling"

    ' -- -ize -> -ise (safe subset) --
    AddSpellingPair usWords, ukWords, pairCount, "organize", "organise"
    AddSpellingPair usWords, ukWords, pairCount, "realize", "realise"
    AddSpellingPair usWords, ukWords, pairCount, "recognize", "recognise"
    AddSpellingPair usWords, ukWords, pairCount, "authorize", "authorise"
    AddSpellingPair usWords, ukWords, pairCount, "characterize", "characterise"
    AddSpellingPair usWords, ukWords, pairCount, "customize", "customise"
    AddSpellingPair usWords, ukWords, pairCount, "emphasize", "emphasise"
    AddSpellingPair usWords, ukWords, pairCount, "finalize", "finalise"
    AddSpellingPair usWords, ukWords, pairCount, "maximize", "maximise"
    AddSpellingPair usWords, ukWords, pairCount, "minimize", "minimise"
    AddSpellingPair usWords, ukWords, pairCount, "normalize", "normalise"
    AddSpellingPair usWords, ukWords, pairCount, "optimize", "optimise"
    AddSpellingPair usWords, ukWords, pairCount, "prioritize", "prioritise"
    AddSpellingPair usWords, ukWords, pairCount, "standardize", "standardise"
    AddSpellingPair usWords, ukWords, pairCount, "summarize", "summarise"
    AddSpellingPair usWords, ukWords, pairCount, "symbolize", "symbolise"
    AddSpellingPair usWords, ukWords, pairCount, "utilize", "utilise"
    AddSpellingPair usWords, ukWords, pairCount, "apologize", "apologise"
    AddSpellingPair usWords, ukWords, pairCount, "capitalize", "capitalise"
    AddSpellingPair usWords, ukWords, pairCount, "criticize", "criticise"
    AddSpellingPair usWords, ukWords, pairCount, "legalize", "legalise"
    AddSpellingPair usWords, ukWords, pairCount, "memorize", "memorise"
    AddSpellingPair usWords, ukWords, pairCount, "patronize", "patronise"
    AddSpellingPair usWords, ukWords, pairCount, "penalize", "penalise"
    AddSpellingPair usWords, ukWords, pairCount, "privatize", "privatise"
    AddSpellingPair usWords, ukWords, pairCount, "harmonize", "harmonise"
    AddSpellingPair usWords, ukWords, pairCount, "economize", "economise"
    AddSpellingPair usWords, ukWords, pairCount, "immunize", "immunise"
    AddSpellingPair usWords, ukWords, pairCount, "neutralize", "neutralise"
    AddSpellingPair usWords, ukWords, pairCount, "stabilize", "stabilise"

    ' -- -ization -> -isation --
    AddSpellingPair usWords, ukWords, pairCount, "organization", "organisation"
    AddSpellingPair usWords, ukWords, pairCount, "authorization", "authorisation"
    AddSpellingPair usWords, ukWords, pairCount, "characterization", "characterisation"
    AddSpellingPair usWords, ukWords, pairCount, "customization", "customisation"
    AddSpellingPair usWords, ukWords, pairCount, "optimization", "optimisation"
    AddSpellingPair usWords, ukWords, pairCount, "normalization", "normalisation"
    AddSpellingPair usWords, ukWords, pairCount, "realization", "realisation"
    AddSpellingPair usWords, ukWords, pairCount, "utilization", "utilisation"
    AddSpellingPair usWords, ukWords, pairCount, "specialization", "specialisation"
    AddSpellingPair usWords, ukWords, pairCount, "globalization", "globalisation"
    AddSpellingPair usWords, ukWords, pairCount, "legalization", "legalisation"
    AddSpellingPair usWords, ukWords, pairCount, "privatization", "privatisation"
    AddSpellingPair usWords, ukWords, pairCount, "harmonization", "harmonisation"
    AddSpellingPair usWords, ukWords, pairCount, "neutralization", "neutralisation"
    AddSpellingPair usWords, ukWords, pairCount, "stabilization", "stabilisation"

    ' -- Safe miscellaneous pairs --
    AddSpellingPair usWords, ukWords, pairCount, "gray", "grey"
    AddSpellingPair usWords, ukWords, pairCount, "plow", "plough"
    AddSpellingPair usWords, ukWords, pairCount, "skeptic", "sceptic"
    AddSpellingPair usWords, ukWords, pairCount, "skeptical", "sceptical"
    AddSpellingPair usWords, ukWords, pairCount, "aluminum", "aluminium"
    AddSpellingPair usWords, ukWords, pairCount, "artifact", "artefact"
    AddSpellingPair usWords, ukWords, pairCount, "aging", "ageing"
    AddSpellingPair usWords, ukWords, pairCount, "pajamas", "pyjamas"
    AddSpellingPair usWords, ukWords, pairCount, "cozy", "cosy"
    AddSpellingPair usWords, ukWords, pairCount, "donut", "doughnut"

    ' Trim arrays to actual size
    ReDim Preserve usWords(0 To pairCount - 1)
    ReDim Preserve ukWords(0 To pairCount - 1)
End Sub

Private Sub AddSpellingPair(ByRef usWords() As String, _
                             ByRef ukWords() As String, _
                             ByRef pairCount As Long, _
                             ByVal usWord As String, _
                             ByVal ukWord As String)
    ' Grow arrays if needed
    If pairCount > UBound(usWords) Then
        ReDim Preserve usWords(0 To UBound(usWords) + 128)
        ReDim Preserve ukWords(0 To UBound(ukWords) + 128)
    End If
    usWords(pairCount) = usWord
    ukWords(pairCount) = ukWord
    pairCount = pairCount + 1
End Sub

' ============================================================
'  PRIVATE: Check if a range is italic
' ============================================================
Private Function IsRangeItalic(rng As Range) As Boolean
    On Error Resume Next
    Dim italicVal As Long
    italicVal = rng.Font.Italic
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsRangeItalic = False
        Exit Function
    End If
    On Error GoTo 0

    ' wdTrue = -1, True = -1; wdUndefined = 9999999 (mixed)
    IsRangeItalic = (italicVal = -1)
End Function

' ============================================================
'  PRIVATE: Check if a range is inside quotation marks
'  Looks at the character immediately before and after the
'  range for smart quotes, straight quotes, or single quotes.
' ============================================================
Private Function IsInsideQuotes(rng As Range, doc As Document) As Boolean
    Dim charBefore As String
    Dim charAfter As String

    On Error Resume Next

    ' Get character before range
    If rng.Start > 0 Then
        charBefore = doc.Range(rng.Start - 1, rng.Start).Text
    Else
        charBefore = ""
    End If
    If Err.Number <> 0 Then
        charBefore = ""
        Err.Clear
    End If

    ' Get character after range
    If rng.End < doc.Content.End Then
        charAfter = doc.Range(rng.End, rng.End + 1).Text
    Else
        charAfter = ""
    End If
    If Err.Number <> 0 Then
        charAfter = ""
        Err.Clear
    End If
    On Error GoTo 0

    ' Check for opening + closing quotes around the word
    ' This catches "word" and 'word' and similar
    If IsOpeningQuote(charBefore) And IsClosingQuote(charAfter) Then
        IsInsideQuotes = True
        Exit Function
    End If

    ' Broader check: scan backward for an unmatched opening quote
    ' within 200 characters
    Dim lookbackStart As Long
    lookbackStart = rng.Start - 200
    If lookbackStart < 0 Then lookbackStart = 0

    On Error Resume Next
    Dim beforeText As String
    beforeText = doc.Range(lookbackStart, rng.Start).Text
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        IsInsideQuotes = False
        Exit Function
    End If
    On Error GoTo 0

    ' Count open vs close quotes in the preceding text
    Dim openCount As Long
    Dim closeCount As Long
    Dim ch As String
    Dim c As Long
    openCount = 0: closeCount = 0
    For c = 1 To Len(beforeText)
        ch = Mid(beforeText, c, 1)
        If IsOpeningQuote(ch) Then openCount = openCount + 1
        If IsClosingQuote(ch) Then closeCount = closeCount + 1
    Next c

    ' If there are more opens than closes, we're inside quotes
    IsInsideQuotes = (openCount > closeCount)
End Function

' ============================================================
'  PRIVATE: Check if a character is an opening quote
' ============================================================
Private Function IsOpeningQuote(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsOpeningQuote = False
        Exit Function
    End If
    Select Case AscW(ch)
        Case 8220  ' left double smart quote "
            IsOpeningQuote = True
        Case 8216  ' left single smart quote '
            IsOpeningQuote = True
        Case Else
            IsOpeningQuote = False
    End Select
End Function

' ============================================================
'  PRIVATE: Check if a character is a closing quote
' ============================================================
Private Function IsClosingQuote(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then
        IsClosingQuote = False
        Exit Function
    End If
    Select Case AscW(ch)
        Case 8221  ' right double smart quote "
            IsClosingQuote = True
        Case 8217  ' right single smart quote '
            IsClosingQuote = True
        Case Else
            IsClosingQuote = False
    End Select
End Function

' ================================================================
' ================================================================
'  RULE 12 -- LICENCE / LICENSE
' ================================================================
' ================================================================

' ============================================================
'  MAIN ENTRY POINT -- Licence/License
' ============================================================
Public Function Check_LicenceLicense(doc As Document) As Collection
    Dim issues As New Collection

    ' Search for both spellings in the document body
    SearchForLicenceIssues doc.Content, doc, issues

    ' Search footnotes
    On Error Resume Next
    Dim fn As Footnote
    For Each fn In doc.Footnotes
        Err.Clear
        SearchForLicenceIssues fn.Range, doc, issues
        If Err.Number <> 0 Then Err.Clear
    Next fn
    On Error GoTo 0

    ' Search endnotes
    On Error Resume Next
    Dim en As Endnote
    For Each en In doc.Endnotes
        Err.Clear
        SearchForLicenceIssues en.Range, doc, issues
        If Err.Number <> 0 Then Err.Clear
    Next en
    On Error GoTo 0

    Set Check_LicenceLicense = issues
End Function

' ============================================================
'  PRIVATE: Search a range for licence/license issues
' ============================================================
Private Sub SearchForLicenceIssues(searchRange As Range, _
                                    doc As Document, _
                                    ByRef issues As Collection)
    Dim searchTerms As Variant
    Dim t As Long

    ' Search for the base forms; skip derivatives that are always correct
    searchTerms = Array("licence", "license", "sub-licence", "sub-license", _
                        "re-licence", "re-license")

    For t = LBound(searchTerms) To UBound(searchTerms)
        SearchSingleLicenceTerm CStr(searchTerms(t)), searchRange, doc, issues
    Next t
End Sub

' ============================================================
'  PRIVATE: Search for a single licence/license term and
'  analyse context
' ============================================================
Private Sub SearchSingleLicenceTerm(ByVal term As String, _
                              searchRange As Range, _
                              doc As Document, _
                              ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String
    Dim contextBefore As String
    Dim contextAfter As String
    Dim wordBefore As String
    Dim wordAfter As String
    Dim issueText As String
    Dim suggestion As String
    Dim usesS As Boolean
    Dim baseIsNoun As Boolean
    Dim baseIsVerb As Boolean

    On Error Resume Next
    Set rng = searchRange.Duplicate
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    With rng.Find
        .ClearFormatting
        .Text = term
        .MatchWholeWord = True
        .MatchCase = False
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Do
        End If
        On Error GoTo 0

        If Not found Then Exit Do

        ' Skip if outside page range
        If Not EngineIsInPageRange(rng) Then
            GoTo ContinueLicenceSearch
        End If

        ' Determine if the found word uses -s- or -c-
        usesS = (InStr(1, LCase(rng.Text), "license") > 0)

        ' Skip "licensed" and "licensing" -- always correct with -s-
        Dim foundLower As String
        foundLower = LCase(Trim(rng.Text))
        If foundLower = "licensed" Or foundLower = "licensing" Then
            GoTo ContinueLicenceSearch
        End If

        ' -- Downgrade italic / quoted text ------------------
        Dim licSeverity As String
        licSeverity = "possible_error"

        If IsRangeItalic(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_LICENCE, locStr, "'" & rng.Text & "' -- in italic text, review manually", "", rng.Start, rng.End, "possible_error")
            issues.Add finding
            GoTo ContinueLicenceSearch
        End If

        If IsInsideQuotes(rng, doc) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_LICENCE, locStr, "'" & rng.Text & "' -- in quoted text, review manually", "", rng.Start, rng.End, "possible_error")
            issues.Add finding
            GoTo ContinueLicenceSearch
        End If

        ' -- Get surrounding context --------------------------
        contextBefore = GetLicenceContextBefore(rng, doc, 50)
        contextAfter = GetLicenceContextAfter(rng, doc, 50)

        ' Extract the last word before the match
        wordBefore = GetLastWordFromContext(contextBefore)

        ' Extract the first word after the match
        wordAfter = GetFirstWordFromContext(contextAfter)

        ' -- Determine noun or verb context -------------------
        baseIsVerb = IsVerbIndicator(wordBefore)
        baseIsNoun = IsNounIndicator(wordBefore) Or IsNounFollower(wordAfter)

        ' -- Decide if there is an finding ----------------------
        issueText = ""
        suggestion = ""

        If usesS And baseIsNoun And Not baseIsVerb Then
            ' "license" used in noun context -- should be "licence"
            issueText = "'" & rng.Text & "' appears in a noun context; " & _
                        "UK convention uses 'licence' for the noun"
            suggestion = ReplaceSWithC(rng.Text)
        ElseIf Not usesS And baseIsVerb And Not baseIsNoun Then
            ' "licence" used in verb context -- should be "license"
            issueText = "'" & rng.Text & "' appears in a verb context; " & _
                        "UK convention uses 'license' for the verb"
            suggestion = ReplaceCWithS(rng.Text)
        ElseIf (usesS And Not baseIsVerb And Not baseIsNoun) Or _
               (Not usesS And Not baseIsVerb And Not baseIsNoun) Then
            ' Context ambiguous
            issueText = "'" & rng.Text & "' -- unable to determine noun/verb context; " & _
                        "review context to ensure correct UK spelling"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        ElseIf usesS And baseIsVerb And baseIsNoun Then
            ' Both indicators present -- ambiguous
            issueText = "'" & rng.Text & "' -- conflicting noun/verb indicators; " & _
                        "review context"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        ElseIf Not usesS And baseIsVerb And baseIsNoun Then
            ' Both indicators present -- ambiguous
            issueText = "'" & rng.Text & "' -- conflicting noun/verb indicators; " & _
                        "review context"
            suggestion = "Review context: 'licence' = noun, 'license' = verb"
        End If

        ' Only create finding if we found something to flag
        If Len(issueText) > 0 Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_LICENCE, locStr, issueText, suggestion, rng.Start, rng.End, "possible_error")
            issues.Add finding
        End If

ContinueLicenceSearch:
        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Do
        End If
        On Error GoTo 0
    Loop
End Sub

' ============================================================
'  PRIVATE: Get text before the match range (up to N chars)
' ============================================================
Private Function GetLicenceContextBefore(rng As Range, doc As Document, _
                                   ByVal charCount As Long) As String
    Dim startPos As Long
    Dim contextRng As Range

    On Error Resume Next
    startPos = rng.Start - charCount
    If startPos < 0 Then startPos = 0

    Set contextRng = doc.Range(startPos, rng.Start)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        GetLicenceContextBefore = ""
        Exit Function
    End If
    On Error GoTo 0

    GetLicenceContextBefore = contextRng.Text
End Function

' ============================================================
'  PRIVATE: Get text after the match range (up to N chars)
' ============================================================
Private Function GetLicenceContextAfter(rng As Range, doc As Document, _
                                  ByVal charCount As Long) As String
    Dim endPos As Long
    Dim contextRng As Range
    Dim docEnd As Long

    On Error Resume Next
    docEnd = doc.Content.End
    endPos = rng.End + charCount
    If endPos > docEnd Then endPos = docEnd

    Set contextRng = doc.Range(rng.End, endPos)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        GetLicenceContextAfter = ""
        Exit Function
    End If
    On Error GoTo 0

    GetLicenceContextAfter = contextRng.Text
End Function

' ============================================================
'  PRIVATE: Extract the last word from a context string
' ============================================================
Private Function GetLastWordFromContext(ByVal text As String) As String
    Dim trimmed As String
    Dim i As Long
    Dim ch As String

    trimmed = Trim(text)
    If Len(trimmed) = 0 Then
        GetLastWordFromContext = ""
        Exit Function
    End If

    ' Walk backward from end to find last word boundary
    For i = Len(trimmed) To 1 Step -1
        ch = Mid(trimmed, i, 1)
        If ch = " " Or ch = vbCr Or ch = vbLf Or ch = vbTab Then
            GetLastWordFromContext = LCase(Mid(trimmed, i + 1))
            Exit Function
        End If
    Next i

    GetLastWordFromContext = LCase(trimmed)
End Function

' ============================================================
'  PRIVATE: Extract the first word from a context string
' ============================================================
Private Function GetFirstWordFromContext(ByVal text As String) As String
    Dim trimmed As String
    Dim spacePos As Long

    trimmed = Trim(text)
    If Len(trimmed) = 0 Then
        GetFirstWordFromContext = ""
        Exit Function
    End If

    spacePos = InStr(1, trimmed, " ")
    If spacePos > 0 Then
        GetFirstWordFromContext = LCase(Left(trimmed, spacePos - 1))
    Else
        GetFirstWordFromContext = LCase(trimmed)
    End If

    ' Strip trailing punctuation
    Dim result As String
    Dim pch As String
    result = GetFirstWordFromContext
    Do While Len(result) > 0
        pch = Right(result, 1)
        If pch Like "[A-Za-z]" Then Exit Do
        result = Left(result, Len(result) - 1)
    Loop
    GetFirstWordFromContext = result
End Function

' ============================================================
'  PRIVATE: Check if a word is a verb indicator
' ============================================================
Private Function IsVerbIndicator(ByVal word As String) As Boolean
    Dim indicators As Variant
    Dim i As Long

    indicators = Array("to", "will", "shall", "may", "must", _
                       "can", "should", "would", "not")

    word = LCase(Trim(word))
    For i = LBound(indicators) To UBound(indicators)
        If word = CStr(indicators(i)) Then
            IsVerbIndicator = True
            Exit Function
        End If
    Next i

    IsVerbIndicator = False
End Function

' ============================================================
'  PRIVATE: Check if a word is a noun indicator
' ============================================================
Private Function IsNounIndicator(ByVal word As String) As Boolean
    Dim indicators As Variant
    Dim i As Long

    indicators = Array("a", "an", "the", "this", "that", "such", _
                       "said", "its", "their", "our", "your", "his", "her")

    word = LCase(Trim(word))
    For i = LBound(indicators) To UBound(indicators)
        If word = CStr(indicators(i)) Then
            IsNounIndicator = True
            Exit Function
        End If
    Next i

    IsNounIndicator = False
End Function

' ============================================================
'  PRIVATE: Check if the word after indicates noun usage
' ============================================================
Private Function IsNounFollower(ByVal word As String) As Boolean
    Dim followers As Variant
    Dim i As Long

    followers = Array("agreement", "holder", "fee", "number", _
                      "plate", "condition")

    word = LCase(Trim(word))
    For i = LBound(followers) To UBound(followers)
        If word = CStr(followers(i)) Then
            IsNounFollower = True
            Exit Function
        End If
    Next i

    IsNounFollower = False
End Function

' ============================================================
'  PRIVATE: Replace -s- with -c- in licence/license words
' ============================================================
Private Function ReplaceSWithC(ByVal word As String) As String
    ReplaceSWithC = Replace(word, "license", "licence", , , vbTextCompare)
    ReplaceSWithC = Replace(ReplaceSWithC, "License", "Licence", , , vbBinaryCompare)
End Function

' ============================================================
'  PRIVATE: Replace -c- with -s- in licence/license words
' ============================================================
Private Function ReplaceCWithS(ByVal word As String) As String
    ReplaceCWithS = Replace(word, "licence", "license", , , vbTextCompare)
    ReplaceCWithS = Replace(ReplaceCWithS, "Licence", "License", , , vbBinaryCompare)
End Function

' ================================================================
' ================================================================
'  RULE 14 -- CHECK / CHEQUE (UK mode only)
'  "check" as a verb (to verify) is valid UK English.
'  Only the financial-instrument noun should be "cheque" in UK.
'  Detects "check" when used as a noun (not a verb) and suggests
'  "cheque". Verb detection uses preceding word context.
' ================================================================
' ================================================================

Public Function Check_CheckCheque(doc As Document) As Collection
    Dim issues As New Collection
    Dim spellingMode As String
    spellingMode = EngineGetSpellingMode()

    ' Only applies in UK mode (US uses "check" for everything)
    If spellingMode <> "UK" Then
        Set Check_CheckCheque = issues
        Exit Function
    End If

    ' Search body text for "check" / "checks" (context-aware)
    SearchCheckCheque doc.Content, doc, issues

    ' Search body text for financial compound phrases
    SearchFinancialCheckCompounds doc.Content, doc, issues

    ' Search footnotes
    On Error Resume Next
    Dim fn As Footnote
    For Each fn In doc.Footnotes
        Err.Clear
        SearchCheckCheque fn.Range, doc, issues
        If Err.Number <> 0 Then Err.Clear
        Err.Clear
        SearchFinancialCheckCompounds fn.Range, doc, issues
        If Err.Number <> 0 Then Err.Clear
    Next fn
    On Error GoTo 0

    Set Check_CheckCheque = issues
End Function

Private Sub SearchCheckCheque(searchRange As Range, doc As Document, _
                               ByRef issues As Collection)
    Dim rng As Range
    Dim foundText As String
    Dim finding As Object
    Dim locStr As String

    ' Search for "check" as whole word
    Dim searchTerms As Variant
    searchTerms = Array("check", "checks")

    Dim si As Long
    For si = LBound(searchTerms) To UBound(searchTerms)
        On Error Resume Next
        Set rng = searchRange.Duplicate
        If Err.Number <> 0 Then Err.Clear: GoTo NextSearchTerm
        On Error GoTo 0

        With rng.Find
            .ClearFormatting
            .Text = CStr(searchTerms(si))
            .MatchWholeWord = True
            .MatchCase = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
        End With

        Dim lastPos As Long
        lastPos = -1
        Do
            On Error Resume Next
            Dim foundIt As Boolean
            foundIt = rng.Find.Execute
            If Err.Number <> 0 Then Err.Clear: Exit Do
            On Error GoTo 0

            If Not foundIt Then Exit Do
            If rng.Start <= lastPos Then Exit Do
            lastPos = rng.Start

            If Not EngineIsInPageRange(rng) Then
                rng.Collapse wdCollapseEnd
                GoTo NextCheckMatch
            End If

            foundText = rng.Text

            ' Determine if this is a verb usage (skip) or noun (flag)
            If IsCheckUsedAsVerb(rng, doc) Then
                rng.Collapse wdCollapseEnd
                GoTo NextCheckMatch
            End If

            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Dim suggestion As String
            If LCase(foundText) = "checks" Then
                suggestion = "cheques"
            Else
                suggestion = "cheque"
            End If

            Set finding = CreateIssueDict(RULE_NAME_CHECK, locStr, _
                "UK spelling: '" & foundText & "' appears to be a noun (financial instrument). Use '" & suggestion & "' in UK English.", _
                suggestion, rng.Start, rng.End, "possible_error")
            issues.Add finding

NextCheckMatch:
            On Error Resume Next
            rng.Collapse wdCollapseEnd
            If Err.Number <> 0 Then Err.Clear: Exit Do
            On Error GoTo 0
        Loop
NextSearchTerm:
    Next si
End Sub

' Determine if "check" is used as a verb by looking at surrounding context.
' Returns True if likely a verb (should NOT be flagged).
Private Function IsCheckUsedAsVerb(rng As Range, doc As Document) As Boolean
    IsCheckUsedAsVerb = False

    ' Get up to 30 chars before the word
    Dim lookStart As Long
    lookStart = rng.Start - 30
    If lookStart < 0 Then lookStart = 0
    Dim beforeText As String
    beforeText = ""
    On Error Resume Next
    If rng.Start > lookStart Then
        beforeText = LCase(doc.Range(lookStart, rng.Start).Text)
    End If
    If Err.Number <> 0 Then beforeText = "": Err.Clear
    On Error GoTo 0

    ' Get up to 20 chars after the word
    Dim afterText As String
    afterText = ""
    Dim lookEnd As Long
    lookEnd = rng.End + 20
    On Error Resume Next
    If lookEnd > doc.Content.End Then lookEnd = doc.Content.End
    If lookEnd > rng.End Then
        afterText = LCase(doc.Range(rng.End, lookEnd).Text)
    End If
    If Err.Number <> 0 Then afterText = "": Err.Clear
    On Error GoTo 0

    ' Extract last word before "check"
    beforeText = Trim(beforeText)
    Dim lastWord As String
    Dim sp As Long
    sp = InStrRev(beforeText, " ")
    If sp > 0 Then
        lastWord = Mid$(beforeText, sp + 1)
    Else
        lastWord = beforeText
    End If

    ' --- Compound prefix check (before "check") ---
    ' Words like "double-check", "cross-check", "spot-check" etc.
    ' are NOT financial. Also handles "double check" (space-separated)
    ' where lastWord = "double" via the space split above.
    Dim compoundPrefixes As Variant
    Dim cp1 As Variant, cp2 As Variant, cp3 As Variant
    cp1 = Array("double", "triple", "quadruple", "spot", _
        "fact", "reality", "re", "counter", "body", "rain")
    cp2 = Array("sound", "spell", "health", "quality", "background", _
        "reference", "security", "safety", "compliance", "system")
    cp3 = Array("gut", "sense", "temperature", "sanity", "mic", _
        "mike", "status", "progress", "wellness", "vibe", _
        "stock", "price", "proof", "ground", "over", _
        "under", "un", "pre", "self")
    Dim vi As Long

    ' Check if lastWord contains a hyphen (e.g. "double-")
    Dim compPrefix As String
    If InStr(lastWord, "-") > 0 Then
        compPrefix = Left$(lastWord, InStr(lastWord, "-") - 1)
    Else
        compPrefix = lastWord
    End If

    ' Check against all compound prefix arrays
    For vi = LBound(cp1) To UBound(cp1)
        If compPrefix = CStr(cp1(vi)) Then
            IsCheckUsedAsVerb = True  ' Not financial
            Exit Function
        End If
    Next vi
    For vi = LBound(cp2) To UBound(cp2)
        If compPrefix = CStr(cp2(vi)) Then
            IsCheckUsedAsVerb = True
            Exit Function
        End If
    Next vi
    For vi = LBound(cp3) To UBound(cp3)
        If compPrefix = CStr(cp3(vi)) Then
            IsCheckUsedAsVerb = True
            Exit Function
        End If
    Next vi

    ' --- Compound suffix check (after "check") ---
    ' Words like "check-in", "check-out", "check-mate" etc.
    ' are NOT financial. "check-book" IS financial (excluded).
    Dim firstCharAfter As String
    afterText = Trim(afterText)
    firstCharAfter = ""
    If Len(afterText) > 0 Then firstCharAfter = Left$(afterText, 1)

    If firstCharAfter = "-" And Len(afterText) > 1 Then
        ' Extract the word after the hyphen
        Dim suffixWord As String
        Dim restAfter As String
        restAfter = Mid$(afterText, 2)
        sp = InStr(1, restAfter, " ")
        If sp > 0 Then
            suffixWord = Left$(restAfter, sp - 1)
        Else
            suffixWord = restAfter
        End If
        suffixWord = LCase(suffixWord)

        Dim nfSuffix1 As Variant, nfSuffix2 As Variant
        nfSuffix1 = Array("in", "out", "up", "list", "mark", "mate", _
            "point", "sum", "box", "off", "room", "er", "ers", "ed", "ing")
        nfSuffix2 = Array("able", "board", "down", "through", "ride", _
            "rein", "bone", "flag", "gate", "land", "line", _
            "pattern", "piece", "rail", "row", "side", "weight", "work")
        ' Note: "book" deliberately absent — check-book IS financial

        For vi = LBound(nfSuffix1) To UBound(nfSuffix1)
            If suffixWord = CStr(nfSuffix1(vi)) Then
                IsCheckUsedAsVerb = True
                Exit Function
            End If
        Next vi
        For vi = LBound(nfSuffix2) To UBound(nfSuffix2)
            If suffixWord = CStr(nfSuffix2(vi)) Then
                IsCheckUsedAsVerb = True
                Exit Function
            End If
        Next vi
    End If

    ' --- Standard verb/noun context analysis ---
    ' Verb indicators: preceded by modal verbs, auxiliaries, etc.
    Dim verbPrecedes As Variant
    verbPrecedes = Array("to", "will", "shall", "must", "should", _
                         "would", "could", "can", "may", "might", _
                         "please", "let", "did", "does", "do", _
                         "not", "always", "also", "then", "and", _
                         "or", "we", "they", "you", "i")
    For vi = LBound(verbPrecedes) To UBound(verbPrecedes)
        If lastWord = CStr(verbPrecedes(vi)) Then
            IsCheckUsedAsVerb = True
            Exit Function
        End If
    Next vi

    ' Verb indicator: followed by certain words
    Dim firstWordAfter As String
    sp = InStr(1, afterText, " ")
    If sp > 0 Then
        firstWordAfter = Left$(afterText, sp - 1)
    Else
        firstWordAfter = afterText
    End If
    ' Strip leading hyphen if present (already handled above for compounds)
    If Left$(firstWordAfter, 1) = "-" Then firstWordAfter = Mid$(firstWordAfter, 2)

    Dim verbFollows As Variant
    verbFollows = Array("that", "whether", "if", "the", "this", _
                        "for", "with", "on", "your", "our", _
                        "his", "her", "its", "their", "my", _
                        "each", "every", "all", "any")
    For vi = LBound(verbFollows) To UBound(verbFollows)
        If firstWordAfter = CStr(verbFollows(vi)) Then
            IsCheckUsedAsVerb = True
            Exit Function
        End If
    Next vi

    ' Noun indicators: preceded by determiners/prepositions
    Dim nounPrecedes As Variant
    nounPrecedes = Array("a", "the", "this", "that", "each", _
                         "every", "your", "our", "his", "her", _
                         "my", "its", "their", "by", "per", _
                         "no", "any", "one", "blank")
    For vi = LBound(nounPrecedes) To UBound(nounPrecedes)
        If lastWord = CStr(nounPrecedes(vi)) Then
            IsCheckUsedAsVerb = False
            Exit Function
        End If
    Next vi

    ' Default: treat as possible noun (flag it as possible_error for review)
    IsCheckUsedAsVerb = False
End Function

' ----------------------------------------------------------------
' Search for financial compound words/phrases containing "check"
' that should use "cheque" in UK English. These are searched as
' literal phrases and flagged unconditionally (no verb/noun analysis).
' ----------------------------------------------------------------
Private Sub SearchFinancialCheckCompounds(searchRange As Range, _
                                          doc As Document, _
                                          ByRef issues As Collection)
    ' Parallel arrays: search terms and their UK suggestions
    ' Split into batches to stay under 25 line-continuation limit
    Dim terms1 As Variant, sugs1 As Variant
    terms1 = Array("checkbook", "check-book", "checkbooks", "check-books", _
        "paycheck", "pay-check", "paychecks", "pay-checks")
    sugs1 = Array("chequebook", "cheque-book", "chequebooks", "cheque-books", _
        "pay cheque", "pay cheque", "pay cheques", "pay cheques")

    Dim terms2 As Variant, sugs2 As Variant
    terms2 = Array("blank check", "blank checks", "bad check", "bad checks", _
        "bounced check", "bounced checks", "rubber check", "rubber checks")
    sugs2 = Array("blank cheque", "blank cheques", "bad cheque", "bad cheques", _
        "bounced cheque", "bounced cheques", "rubber cheque", "rubber cheques")

    Dim terms3 As Variant, sugs3 As Variant
    terms3 = Array("cancelled check", "canceled check", _
        "certified check", "certified checks", _
        "cashier's check", "cashiers check")
    sugs3 = Array("cancelled cheque", "cancelled cheque", _
        "certified cheque", "certified cheques", _
        "cashier's cheque", "cashier's cheque")

    Dim terms4 As Variant, sugs4 As Variant
    terms4 = Array("traveller's check", "traveler's check", _
        "travellers check", "travelers check", _
        "traveller's checks", "traveler's checks")
    sugs4 = Array("traveller's cheque", "traveller's cheque", _
        "travellers' cheque", "travellers' cheque", _
        "traveller's cheques", "traveller's cheques")

    Dim terms5 As Variant, sugs5 As Variant
    terms5 = Array("travellers checks", "travelers checks", _
        "personal check", "personal checks", _
        "bank check", "bank checks")
    sugs5 = Array("travellers' cheques", "travellers' cheques", _
        "personal cheque", "personal cheques", _
        "bank cheque", "bank cheques")

    Dim terms6 As Variant, sugs6 As Variant
    terms6 = Array("post-dated check", "postdated check", _
        "stale check", "stale checks", _
        "dishonoured check", "dishonored check")
    sugs6 = Array("post-dated cheque", "post-dated cheque", _
        "stale cheque", "stale cheques", _
        "dishonoured cheque", "dishonoured cheque")

    Dim terms7 As Variant, sugs7 As Variant
    terms7 = Array("check stub", "check stubs", "check fraud", _
        "check forgery", "check clearing", _
        "check guarantee", "check number", "check numbers")
    sugs7 = Array("cheque stub", "cheque stubs", "cheque fraud", _
        "cheque forgery", "cheque clearing", _
        "cheque guarantee", "cheque number", "cheque numbers")

    ' Process each batch
    SearchFinancialBatch searchRange, doc, issues, terms1, sugs1, True
    SearchFinancialBatch searchRange, doc, issues, terms2, sugs2, False
    SearchFinancialBatch searchRange, doc, issues, terms3, sugs3, False
    SearchFinancialBatch searchRange, doc, issues, terms4, sugs4, False
    SearchFinancialBatch searchRange, doc, issues, terms5, sugs5, False
    SearchFinancialBatch searchRange, doc, issues, terms6, sugs6, False
    SearchFinancialBatch searchRange, doc, issues, terms7, sugs7, False
End Sub

Private Sub SearchFinancialBatch(searchRange As Range, _
                                  doc As Document, _
                                  ByRef issues As Collection, _
                                  terms As Variant, _
                                  suggestions As Variant, _
                                  wholeWord As Boolean)
    Dim ti As Long
    Dim rng As Range
    Dim finding As Object
    Dim locStr As String

    For ti = LBound(terms) To UBound(terms)
        On Error Resume Next
        Set rng = searchRange.Duplicate
        If Err.Number <> 0 Then Err.Clear: GoTo NextFinTerm
        On Error GoTo 0

        With rng.Find
            .ClearFormatting
            .Text = CStr(terms(ti))
            .MatchWholeWord = wholeWord
            .MatchCase = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
        End With

        Dim lastPos As Long
        lastPos = -1
        Do
            On Error Resume Next
            Dim foundIt As Boolean
            foundIt = rng.Find.Execute
            If Err.Number <> 0 Then Err.Clear: Exit Do
            On Error GoTo 0

            If Not foundIt Then Exit Do
            If rng.Start <= lastPos Then Exit Do
            lastPos = rng.Start

            If Not EngineIsInPageRange(rng) Then
                rng.Collapse wdCollapseEnd
                GoTo NextFinMatch
            End If

            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_CHECK, locStr, _
                "UK spelling: '" & rng.Text & "' should be '" & _
                CStr(suggestions(ti)) & "' in UK English.", _
                CStr(suggestions(ti)), rng.Start, rng.End, "possible_error", True)
            issues.Add finding

NextFinMatch:
            On Error Resume Next
            rng.Collapse wdCollapseEnd
            If Err.Number <> 0 Then Err.Clear: Exit Do
            On Error GoTo 0
        Loop
NextFinTerm:
    Next ti
End Sub

' ================================================================
' ================================================================
'  RULE 13 -- COLOUR FORMATTING
' ================================================================
' ================================================================

' ============================================================
'  MAIN ENTRY POINT -- Colour Formatting
' ============================================================
Public Function Check_ColourFormatting(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraColor As Long
    Dim colourCounts As Object
    Dim dominantColour As Long
    Dim maxCount As Long

    Const WD_COLOR_AUTOMATIC As Long = -16777216

    ' -- Build hyperlink position set once (avoid O(n^2)) ------
    Dim hlStarts As Object, hlEnds As Object
    Set hlStarts = CreateObject("Scripting.Dictionary")
    Set hlEnds = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim hl As Hyperlink
    Dim hlIdx As Long: hlIdx = 0
    For Each hl In doc.Hyperlinks
        Err.Clear
        hlStarts.Add hlIdx, hl.Range.Start
        hlEnds.Add hlIdx, hl.Range.End
        If Err.Number <> 0 Then Err.Clear
        hlIdx = hlIdx + 1
    Next hl
    On Error GoTo 0

    ' -- Pass 1: count paragraph-level colours -----------------
    Set colourCounts = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextPC1

        If Not EngineIsInPageRange(paraRange) Then GoTo NextPC1

        paraColor = paraRange.Font.Color
        If Err.Number <> 0 Then Err.Clear: GoTo NextPC1

        ' Skip indeterminate (mixed-colour paragraphs counted in pass 2)
        If paraColor = 9999999 Then GoTo NextPC1

        If colourCounts.Exists(paraColor) Then
            colourCounts(paraColor) = colourCounts(paraColor) + 1
        Else
            colourCounts.Add paraColor, 1
        End If
NextPC1:
    Next para
    On Error GoTo 0

    ' -- Determine dominant colour -----------------------------
    If colourCounts.Count = 0 Then
        Set Check_ColourFormatting = issues
        Exit Function
    End If

    dominantColour = 0: maxCount = 0
    Dim colourKey As Variant
    For Each colourKey In colourCounts.keys
        If colourCounts(colourKey) > maxCount Then
            maxCount = colourCounts(colourKey)
            dominantColour = CLng(colourKey)
        End If
    Next colourKey

    ' -- Pass 2: flag paragraphs with non-standard colours -----
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextPC2

        If Not EngineIsInPageRange(paraRange) Then GoTo NextPC2

        ' Skip heading-styled paragraphs
        Dim styleName As String
        styleName = ""
        styleName = para.Style.NameLocal
        If Err.Number <> 0 Then Err.Clear: styleName = ""
        If LCase(Left(styleName, 7)) = "heading" Then GoTo NextPC2

        paraColor = paraRange.Font.Color
        If Err.Number <> 0 Then Err.Clear: GoTo NextPC2

        ' Skip dominant, automatic, or indeterminate
        If paraColor = dominantColour Or _
           paraColor = WD_COLOR_AUTOMATIC Or _
           paraColor = 9999999 Then GoTo NextPC2

        ' Skip if inside a hyperlink
        If IsRangeInsideHyperlink(paraRange, hlStarts, hlEnds) Then GoTo NextPC2

        ' Flag this paragraph
        FlushColourGroup doc, issues, paraRange.Start, paraRange.End, paraColor

NextPC2:
    Next para
    On Error GoTo 0

    Set Check_ColourFormatting = issues
End Function

' ============================================================
'  PRIVATE: Flush a grouped colour finding
' ============================================================
Private Sub FlushColourGroup(doc As Document, _
                              ByRef issues As Collection, _
                              ByVal startPos As Long, _
                              ByVal endPos As Long, _
                              ByVal fontColor As Long)
    Dim finding As Object
    Dim locStr As String
    Dim hexStr As String
    Dim rng As Range

    hexStr = ColourToHex(fontColor)

    On Error Resume Next
    Set rng = doc.Range(startPos, endPos)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    locStr = EngineGetLocationString(rng, doc)
    If Err.Number <> 0 Then
        locStr = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0

    Dim previewText As String
    On Error Resume Next
    previewText = Left(rng.Text, 60)
    If Err.Number <> 0 Then
        previewText = "(text unavailable)"
        Err.Clear
    End If
    On Error GoTo 0

    Set finding = CreateIssueDict(RULE_NAME_COLOUR, locStr, "Non-standard font colour " & hexStr & " detected: '" & previewText & "'", "Change font colour to match document default", startPos, endPos, "possible_error")
    issues.Add finding
End Sub

' ============================================================
'  PRIVATE: Convert a Long colour value to hex string
' ============================================================
Private Function ColourToHex(ByVal colorVal As Long) As String
    Dim cR As Long
    Dim cG As Long
    Dim cB As Long

    ' Word stores colours as BGR in Long format
    cR = colorVal Mod 256
    cG = (colorVal \ 256) Mod 256
    cB = (colorVal \ 65536) Mod 256

    ColourToHex = "#" & Right("0" & Hex(cR), 2) & _
                        Right("0" & Hex(cG), 2) & _
                        Right("0" & Hex(cB), 2)
End Function

' ============================================================
'  PRIVATE: Check if a run is inside a hyperlink
' ============================================================
Private Function IsRangeInsideHyperlink(rng As Range, _
                                        hlStarts As Object, _
                                        hlEnds As Object) As Boolean
    Dim i As Long
    For i = 0 To hlStarts.Count - 1
        If hlStarts(i) <= rng.Start And hlEnds(i) >= rng.End Then
            IsRangeInsideHyperlink = True
            Exit Function
        End If
    Next i
    IsRangeInsideHyperlink = False
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsWhitelistedTerm
' ----------------------------------------------------------------
Private Function EngineIsWhitelistedTerm(ByVal term As String) As Boolean
    On Error Resume Next
    EngineIsWhitelistedTerm = Application.Run("PleadingsEngine.IsWhitelistedTerm", term)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsWhitelistedTerm: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsWhitelistedTerm = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetSpellingMode
' ----------------------------------------------------------------
Private Function EngineGetSpellingMode() As String
    On Error Resume Next
    EngineGetSpellingMode = Application.Run("PleadingsEngine.GetSpellingMode")
    If Err.Number <> 0 Then
        Debug.Print "EngineGetSpellingMode: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetSpellingMode = "UK"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_Terms.bas

```vb
Attribute VB_Name = "Rules_Terms"
' ============================================================
' Rules_Terms.bas
' Combined module for term-related rules:
'   Rule05 - Custom term whitelist
'   Rule07 - Defined terms
'
' RETIRED (not engine-wired):
'   Rule23 - Phrase consistency: kept for backwards compatibility
'     but not dispatched by RunAllPleadingsRules. Retired due to
'     high false-positive rate on common legal phrases.
' ============================================================
Option Explicit

Private Const RULE05_NAME As String = "custom_term_whitelist"
Private Const RULE07_NAME As String = "defined_terms"
' RETIRED -- NOT ENGINE-WIRED: phrase_consistency kept only for backwards compat
Private Const RETIRED_RETIRED_RULE23_NAME As String = "phrase_consistency"

' ============================================================
'  PRIVATE HELPERS (Rule07)
' ============================================================

' -- Helper: check if quoted text looks like a sentence/quote --
' rather than a defined term (questions, long phrases, etc.)
Private Function LooksLikeSentence(ByVal txt As String) As Boolean
    LooksLikeSentence = False

    ' Contains a question mark -- it's a question, not a term
    If InStr(1, txt, "?") > 0 Then
        LooksLikeSentence = True
        Exit Function
    End If

    ' Very long text is unlikely to be a term name
    If Len(txt) > 60 Then
        LooksLikeSentence = True
        Exit Function
    End If

    ' Count spaces to estimate word count
    Dim spaceCount As Long
    Dim ci As Long
    spaceCount = 0
    For ci = 1 To Len(txt)
        If Mid$(txt, ci, 1) = " " Then spaceCount = spaceCount + 1
    Next ci

    ' More than 8 words is almost certainly a sentence
    If spaceCount >= 8 Then
        LooksLikeSentence = True
        Exit Function
    End If
End Function

' -- Helper: remove hyphens from a term ----------------------
Private Function RemoveHyphens(ByVal term As String) As String
    RemoveHyphens = Replace(term, "-", "")
End Function

' -- Helper: count occurrences of a term in document text ----
Private Function CountTermInDoc(doc As Document, ByVal searchTerm As String) As Long
    Dim rng As Range
    Dim cnt As Long
    cnt = 0

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = searchTerm
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False
    End With

    Dim lastPos As Long
    lastPos = -1
    Do While rng.Find.Execute
        If rng.Start <= lastPos Then Exit Do  ' stall guard
        lastPos = rng.Start
        cnt = cnt + 1
        rng.Collapse wdCollapseEnd
    Loop

    CountTermInDoc = cnt
End Function

' -- Helper: find first occurrence of a term and return range -
Private Function FindTermRange(doc As Document, ByVal searchTerm As String, _
                                matchCase As Boolean) As Range
    Dim rng As Range
    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = searchTerm
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = matchCase
        .MatchWholeWord = True
        .MatchWildcards = False
    End With

    If rng.Find.Execute Then
        Set FindTermRange = rng
    Else
        Set FindTermRange = Nothing
    End If
End Function

' ============================================================
'  PRIVATE HELPERS (Rule23)
' ============================================================

' -- Determine if a phrase is a single word (no spaces) ------
Private Function IsSingleWord(ByVal phrase As String) As Boolean
    IsSingleWord = (InStr(1, phrase, " ") = 0)
End Function

' -- Check a single phrase group for consistency -------------
'  Counts each phrase, determines dominant, flags minorities.
' ------------------------------------------------------------
Private Sub CheckPhraseGroup(doc As Document, _
                              phrases As Variant, _
                              ByRef issues As Collection)
    Dim counts() As Long
    Dim phraseCount As Long
    Dim p As Long
    Dim dominantIdx As Long
    Dim dominantCount As Long
    Dim usedCount As Long

    phraseCount = UBound(phrases) - LBound(phrases) + 1
    ReDim counts(LBound(phrases) To UBound(phrases))

    ' -- Count occurrences of each phrase ---------------------
    For p = LBound(phrases) To UBound(phrases)
        counts(p) = CountPhrase(doc, CStr(phrases(p)))
    Next p

    ' -- Determine how many phrases in this group are used ----
    usedCount = 0
    dominantIdx = LBound(phrases)
    dominantCount = counts(LBound(phrases))

    For p = LBound(phrases) To UBound(phrases)
        If counts(p) > 0 Then usedCount = usedCount + 1
        If counts(p) > dominantCount Then
            dominantCount = counts(p)
            dominantIdx = p
        End If
    Next p

    ' Only flag if more than one phrase in the group is used
    If usedCount < 2 Then Exit Sub

    ' -- Flag all minority phrase occurrences -----------------
    For p = LBound(phrases) To UBound(phrases)
        If counts(p) > 0 And p <> dominantIdx Then
            FlagPhraseOccurrences doc, CStr(phrases(p)), CStr(phrases(dominantIdx)), issues
        End If
    Next p
End Sub

' -- Count occurrences of a phrase in the document -----------
'  Single words use MatchWholeWord=True for proper boundaries.
'  Multi-word phrases use MatchWholeWord=False with manual
'  word-boundary validation at both ends to prevent matching
'  fragments inside larger words.
' ------------------------------------------------------------
Private Function CountPhrase(doc As Document, phrase As String) As Long
    Dim rng As Range
    Dim cnt As Long
    Dim found As Boolean
    Dim singleWord As Boolean

    cnt = 0
    singleWord = IsSingleWord(phrase)

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = phrase
        .MatchWholeWord = singleWord    ' True for single words
        .MatchCase = False
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastMatchStart As Long
    lastMatchStart = -1

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do

        ' Stall guard
        If rng.Start <= lastMatchStart Then Exit Do
        lastMatchStart = rng.Start

        ' Verify word boundaries: even for single words, MatchWholeWord
        ' can behave inconsistently with hyphens and special chars.
        If Not IsWordBoundaryMatch(rng, doc) Then
            On Error Resume Next
            rng.Collapse wdCollapseEnd
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
            On Error GoTo 0
            GoTo NextCountMatch
        End If

        If EngineIsInPageRange(rng) Then
            cnt = cnt + 1
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
NextCountMatch:
    Loop

    CountPhrase = cnt
End Function

' -- Flag all occurrences of a minority phrase ---------------
'  Uses same boundary logic as CountPhrase.
Private Sub FlagPhraseOccurrences(doc As Document, _
                                   minorityPhrase As String, _
                                   dominantPhrase As String, _
                                   ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
    Dim locStr As String
    Dim singleWord As Boolean

    singleWord = IsSingleWord(minorityPhrase)

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = minorityPhrase
        .MatchWholeWord = singleWord    ' True for single words
        .MatchCase = False
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim lastMatchStart As Long
    lastMatchStart = -1

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do

        ' Stall guard
        If rng.Start <= lastMatchStart Then Exit Do
        lastMatchStart = rng.Start

        ' Verify word boundaries: even for single words, MatchWholeWord
        ' can behave inconsistently with hyphens and special chars.
        If Not IsWordBoundaryMatch(rng, doc) Then
            On Error Resume Next
            rng.Collapse wdCollapseEnd
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
            On Error GoTo 0
            GoTo NextFlagMatch
        End If

        If EngineIsInPageRange(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RETIRED_RULE23_NAME, locStr, "Inconsistent phrase: '" & rng.Text & "' used", "Use '" & dominantPhrase & "' for consistency (dominant style)", rng.Start, rng.End, "error")
            issues.Add finding
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
NextFlagMatch:
    Loop
End Sub

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

' ============================================================
'  RULE 07: DEFINED TERMS
' ============================================================
Public Function Check_DefinedTerms(doc As Document) As Collection
    Dim issues As New Collection

    ' Dictionary: term (String) -> Array(definitionParaIdx, rangeStart, rangeEnd)
    Dim definedTerms As Object
    Set definedTerms = CreateObject("Scripting.Dictionary")
    Dim defInfo() As Variant
    Dim mInfo() As Variant
    Dim hInfo() As Variant
    Dim pInfo() As Variant
    Dim rng As Range
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim paraText As String

    ' ==========================================================
    '  PASS 1: Scan for defined terms
    ' ==========================================================

    ' -- Pattern A: Smart-quoted defined terms ----------------
    ' Each pattern section uses scoped OERN around its Word OM calls.
    ' Use quote preference from engine to determine which quotes to search
    Dim leftSmart As String
    Dim rightSmart As String
    Dim termQPref As String
    termQPref = EngineGetTermQuotePref()
    If termQPref = "SINGLE" Then
        leftSmart = ChrW(8216)   ' left single smart quote
        rightSmart = ChrW(8217)  ' right single smart quote
    Else
        leftSmart = ChrW(8220)   ' left double smart quote
        rightSmart = ChrW(8221)  ' right double smart quote
    End If

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = leftSmart & "[A-Z]"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWildcards = True
    End With

    Dim lastPos As Long
    lastPos = -1
    On Error Resume Next
    Do While rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: Exit Do
        If rng.Start <= lastPos Then Exit Do  ' stall guard
        lastPos = rng.Start
        If Not EngineIsInPageRange(rng) Then
            rng.Collapse wdCollapseEnd
            GoTo NextSmartFind
        End If

        ' Expand to find the closing smart quote
        Dim startPos As Long
        startPos = rng.Start
        Dim expandedRng As Range
        Set expandedRng = doc.Range(startPos, startPos)

        ' Search forward for closing smart quote (max 100 chars)
        Dim endSearch As Long
        endSearch = startPos + 100
        If endSearch > doc.Content.End Then endSearch = doc.Content.End
        Set expandedRng = doc.Range(startPos, endSearch)
        Dim fullText As String
        fullText = expandedRng.Text

        Dim closePos As Long
        closePos = InStr(2, fullText, rightSmart)
        If closePos > 1 Then
            Dim termText As String
            ' Extract between quotes (skip the opening quote)
            termText = Mid$(fullText, 2, closePos - 2)
            ' Skip if it looks like a sentence/quote rather than a defined term
            ' (too long, contains question marks, or has many words)
            If Len(Trim$(termText)) > 0 And Not definedTerms.Exists(termText) _
               And Not LooksLikeSentence(termText) Then
                ReDim defInfo(0 To 2)
                defInfo(0) = 0 ' paragraph index (approximate)
                defInfo(1) = startPos  ' range start (includes opening quote)
                defInfo(2) = startPos + CLng(closePos)  ' range end (includes closing quote)
                definedTerms.Add termText, defInfo
            End If
        End If

        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: Exit Do
NextSmartFind:
    Loop
    On Error GoTo 0

    ' -- Pattern A2: Straight-quoted defined terms ("-quoted) --
    ' Mirrors Pattern A but for straight double quotes
    Dim straightQ As String
    straightQ = Chr$(34)  ' straight double quote "

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = straightQ & "[A-Z]"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWildcards = True
    End With

    lastPos = -1
    On Error Resume Next
    Do While rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: Exit Do
        If rng.Start <= lastPos Then Exit Do  ' stall guard
        lastPos = rng.Start
        If Not EngineIsInPageRange(rng) Then
            rng.Collapse wdCollapseEnd
            GoTo NextStraightFind
        End If

        startPos = rng.Start
        endSearch = startPos + 100
        If endSearch > doc.Content.End Then endSearch = doc.Content.End
        Set expandedRng = doc.Range(startPos, endSearch)
        fullText = expandedRng.Text

        closePos = InStr(2, fullText, straightQ)
        If closePos > 1 Then
            Dim sqTermText As String
            sqTermText = Mid$(fullText, 2, closePos - 2)
            If Len(Trim$(sqTermText)) > 0 And Not definedTerms.Exists(sqTermText) _
               And Not LooksLikeSentence(sqTermText) Then
                Dim sqInfo() As Variant
                ReDim sqInfo(0 To 2)
                sqInfo(0) = 0
                sqInfo(1) = startPos
                sqInfo(2) = startPos + CLng(closePos)
                definedTerms.Add sqTermText, sqInfo
            End If
        End If

        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: Exit Do
NextStraightFind:
    Loop
    On Error GoTo 0

    ' -- Pattern B: "X means " or "X has the meaning " -------
    paraIdx = 0
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        paraIdx = paraIdx + 1
        If Not EngineIsInPageRange(para.Range) Then GoTo NextParaMeans

        paraText = para.Range.Text
        Dim meansPos As Long
        meansPos = InStr(1, paraText, " means ", vbTextCompare)
        If meansPos > 1 Then
            ' Extract term before " means "
            Dim beforeMeans As String
            beforeMeans = Trim$(Left$(paraText, meansPos - 1))
            ' Take last quoted or significant phrase
            Dim lastQuoteStart As Long
            lastQuoteStart = InStrRev(beforeMeans, leftSmart)
            If lastQuoteStart = 0 Then lastQuoteStart = InStrRev(beforeMeans, """")
            If lastQuoteStart > 0 Then
                Dim afterQuote As String
                afterQuote = Mid$(beforeMeans, lastQuoteStart + 1)
                Dim endQuote As Long
                endQuote = InStr(1, afterQuote, rightSmart)
                If endQuote = 0 Then endQuote = InStr(1, afterQuote, """")
                If endQuote > 1 Then
                    Dim meansTerm As String
                    meansTerm = Left$(afterQuote, endQuote - 1)
                    If Len(meansTerm) > 0 And Not definedTerms.Exists(meansTerm) Then
                        ReDim mInfo(0 To 2)
                        mInfo(0) = paraIdx
                        mInfo(1) = para.Range.Start
                        mInfo(2) = para.Range.Start + meansPos
                        definedTerms.Add meansTerm, mInfo
                    End If
                End If
            End If
        End If

        ' Check for "has the meaning"
        Dim htmPos As Long
        htmPos = InStr(1, paraText, " has the meaning ", vbTextCompare)
        If htmPos > 1 Then
            Dim beforeHTM As String
            beforeHTM = Trim$(Left$(paraText, htmPos - 1))
            lastQuoteStart = InStrRev(beforeHTM, leftSmart)
            If lastQuoteStart = 0 Then lastQuoteStart = InStrRev(beforeHTM, """")
            If lastQuoteStart > 0 Then
                afterQuote = Mid$(beforeHTM, lastQuoteStart + 1)
                endQuote = InStr(1, afterQuote, rightSmart)
                If endQuote = 0 Then endQuote = InStr(1, afterQuote, """")
                If endQuote > 1 Then
                    Dim htmTerm As String
                    htmTerm = Left$(afterQuote, endQuote - 1)
                    If Len(htmTerm) > 0 And Not definedTerms.Exists(htmTerm) Then
                        ReDim hInfo(0 To 2)
                        hInfo(0) = paraIdx
                        hInfo(1) = para.Range.Start
                        hInfo(2) = para.Range.Start + htmPos
                        definedTerms.Add htmTerm, hInfo
                    End If
                End If
            End If
        End If
NextParaMeans:
    Next para
    On Error GoTo 0

    ' -- Pattern C: Parenthetical definitions (the "Term") ---
    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = "(the " & leftSmart & "[A-Z]"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWildcards = True
    End With

    lastPos = -1
    On Error Resume Next
    Do While rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: Exit Do
        If rng.Start <= lastPos Then Exit Do  ' stall guard
        lastPos = rng.Start
        If Not EngineIsInPageRange(rng) Then
            rng.Collapse wdCollapseEnd
            GoTo NextParenFind
        End If

        startPos = rng.Start
        endSearch = startPos + 120
        If endSearch > doc.Content.End Then endSearch = doc.Content.End
        Set expandedRng = doc.Range(startPos, endSearch)
        fullText = expandedRng.Text

        ' Find closing smart quote then closing paren
        closePos = InStr(6, fullText, rightSmart)
        If closePos > 6 Then
            ' Extract between the smart quotes
            Dim pOpenQ As Long
            pOpenQ = InStr(1, fullText, leftSmart)
            If pOpenQ > 0 Then
                Dim parenTerm As String
                parenTerm = Mid$(fullText, pOpenQ + 1, closePos - pOpenQ - 1)
                If Len(parenTerm) > 0 And Not definedTerms.Exists(parenTerm) Then
                    ReDim pInfo(0 To 2)
                    pInfo(0) = 0
                    pInfo(1) = startPos + pOpenQ
                    pInfo(2) = startPos + CLng(closePos)
                    definedTerms.Add parenTerm, pInfo
                End If
            End If
        End If

        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: Exit Do
NextParenFind:
    Loop
    On Error GoTo 0

    ' ==========================================================
    '  PASS 2: Validate each defined term
    ' ==========================================================
    Dim termKey As Variant
    For Each termKey In definedTerms.keys
        Dim term As String
        term = CStr(termKey)
        Dim tInfo As Variant
        tInfo = definedTerms(termKey)

        ' -- Check A: Lowercase variant (inconsistent capitalisation) --
        '    Flag ALL occurrences (not just the first)
        Dim lcTerm2 As String
        lcTerm2 = LCase(Left$(term, 1)) & Mid$(term, 2)
        If lcTerm2 <> term Then
            On Error Resume Next
            Dim lcRng As Range
            Set lcRng = doc.Content.Duplicate
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo SkipLCCheck
            On Error GoTo 0
            With lcRng.Find
                .ClearFormatting
                .Text = lcTerm2
                .Forward = True
                .Wrap = wdFindStop
                .MatchCase = True
                .MatchWholeWord = True
                .MatchWildcards = False
            End With
            Dim lcLastPos As Long
            lcLastPos = -1
            On Error Resume Next
            Do While lcRng.Find.Execute
                If Err.Number <> 0 Then Err.Clear: Exit Do
                If lcRng.Start <= lcLastPos Then Exit Do
                lcLastPos = lcRng.Start
                If EngineIsInPageRange(lcRng) Then
                    Dim findingLC As Object
                    Dim locLC As String
                    Err.Clear
                    locLC = EngineGetLocationString(lcRng, doc)
                    If Err.Number <> 0 Then locLC = "unknown location": Err.Clear
                    Set findingLC = CreateIssueDict(RULE07_NAME, locLC, "Inconsistent capitalisation: '" & lcTerm2 & "' found but '" & term & "' is the defined term", "Use '" & term & "' consistently", lcRng.Start, lcRng.End, "error")
                    issues.Add findingLC
                End If
                lcRng.Collapse wdCollapseEnd
                If Err.Number <> 0 Then Err.Clear: Exit Do
            Loop
            On Error GoTo 0
        End If
SkipLCCheck:

        ' -- Check B: Hyphenated/unhyphenated variant ----------
        If InStr(1, term, "-") > 0 Then
            Dim noHyphen As String
            noHyphen = RemoveHyphens(term)
            Dim nhRng As Range
            On Error Resume Next
            Set nhRng = FindTermRange(doc, noHyphen, False)
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo SkipHyphenCheck
            On Error GoTo 0
            If Not nhRng Is Nothing Then
                If EngineIsInPageRange(nhRng) Then
                    Dim findingH As Object
                    Dim locH As String
                    On Error Resume Next
                    locH = EngineGetLocationString(nhRng, doc)
                    If Err.Number <> 0 Then locH = "unknown location": Err.Clear
                    On Error GoTo 0
                    Set findingH = CreateIssueDict(RULE07_NAME, locH, "Hyphenation variant: '" & noHyphen & "' found but defined term uses hyphen: '" & term & "'", "Use the defined form: '" & term & "'", nhRng.Start, nhRng.End, "error")
                    issues.Add findingH
                End If
            End If
        Else
            ' Term has no hyphen -- check if hyphenated variant exists
            ' Try common hyphenation points (before common prefixes)
            ' This is a best-effort check
        End If
SkipHyphenCheck:

        ' -- Check C: Defined term never referenced ------------
        Dim totalCount As Long
        totalCount = CountTermInDoc(doc, term)
        If totalCount <= 1 Then
            ' Only appears at the definition site
            Dim findingUnused As Object
            Dim unusedRng As Range
            On Error Resume Next
            Set unusedRng = doc.Range(CLng(tInfo(1)), CLng(tInfo(2)))
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo SkipUnusedCheck
            Dim locUnused As String
            locUnused = EngineGetLocationString(unusedRng, doc)
            If Err.Number <> 0 Then locUnused = "unknown location": Err.Clear
            On Error GoTo 0
            Set findingUnused = CreateIssueDict(RULE07_NAME, locUnused, "Defined term never referenced: '" & term & "' is defined but not used elsewhere in the document.", "", CLng(tInfo(1)), CLng(tInfo(2)), "possible_error")
            issues.Add findingUnused
        End If
SkipUnusedCheck:
    Next termKey

    Set Check_DefinedTerms = issues
End Function

' ============================================================
'  RETIRED RULE 23: PHRASE CONSISTENCY
'  NOT dispatched by RunAllPleadingsRules. Retired due to high
'  false-positive rate on common legal phrases.
'  Kept ONLY for backwards compatibility if called externally.
'  Will emit a debug warning when invoked.
' ============================================================
Public Function Check_PhraseConsistency(doc As Document) As Collection
    Debug.Print "WARNING: Rules_Terms.Check_PhraseConsistency is RETIRED (Rule23). " & _
                "Not dispatched by RunAllPleadingsRules."
    Dim issues As New Collection

    ' -- Define phrase groups ---------------------------------
    ' Each group is an array of synonymous phrases
    Dim groups(0 To 9) As Variant

    groups(0) = Array("not later than", "no later than")
    groups(1) = Array("in respect of", "with respect to", "in relation to")
    groups(2) = Array("pursuant to", "in accordance with")
    groups(3) = Array("notwithstanding", "despite", "regardless of")
    groups(4) = Array("prior to", "before")
    groups(5) = Array("subsequent to", "after", "following")
    groups(6) = Array("in the event that", "if", "where")
    groups(7) = Array("save that", "except that", "provided that")
    groups(8) = Array("forthwith", "immediately", "without delay")
    groups(9) = Array("hereby", "by this")

    ' -- Process each group -----------------------------------
    Dim g As Long
    For g = 0 To 9
        CheckPhraseGroup doc, groups(g), issues
    Next g

    Set Check_PhraseConsistency = issues
End Function


' ----------------------------------------------------------------
'  PRIVATE: Check that a Find match sits on word boundaries
'  (char before/after is not a letter/digit). Prevents partial
'  matches inside larger words or compounds.
'  Used for multi-word phrases where MatchWholeWord is False.
' ----------------------------------------------------------------
Private Function IsWordBoundaryMatch(rng As Range, doc As Document) As Boolean
    IsWordBoundaryMatch = True
    On Error Resume Next
    ' Check character before match
    If rng.Start > 0 Then
        Dim bRng As Range
        Set bRng = doc.Range(rng.Start - 1, rng.Start)
        If Err.Number = 0 Then
            Dim bc As String
            bc = bRng.Text
            If Err.Number = 0 Then
                If IsWordChar(bc) Then
                    IsWordBoundaryMatch = False
                    Err.Clear: On Error GoTo 0
                    Exit Function
                End If
            Else
                Err.Clear
            End If
        Else
            Err.Clear
        End If
    End If
    ' Check character after match
    If rng.End < doc.Content.End Then
        Dim aRng As Range
        Set aRng = doc.Range(rng.End, rng.End + 1)
        If Err.Number = 0 Then
            Dim ac As String
            ac = aRng.Text
            If Err.Number = 0 Then
                If IsWordChar(ac) Then
                    IsWordBoundaryMatch = False
                    Err.Clear: On Error GoTo 0
                    Exit Function
                End If
            Else
                Err.Clear
            End If
        Else
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Word character test: letters (A-Z, a-z), digits, and extended
'  Latin characters. Used for boundary checking.
' ----------------------------------------------------------------
Private Function IsWordChar(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsWordChar = (ch >= "A" And ch <= "Z") Or _
                 (ch >= "a" And ch <= "z") Or _
                 (ch >= "0" And ch <= "9") Or _
                 (AscW(ch) >= 192 And AscW(ch) <= 687)
End Function

' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: EngineSetWhitelist
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
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

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetTermQuotePref
' ----------------------------------------------------------------
Private Function EngineGetTermQuotePref() As String
    On Error Resume Next
    EngineGetTermQuotePref = Application.Run("PleadingsEngine.GetTermQuotePref")
    If Err.Number <> 0 Then
        Debug.Print "EngineGetTermQuotePref: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetTermQuotePref = "DOUBLE"
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetTermFormatPref
' ----------------------------------------------------------------
Private Function EngineGetTermFormatPref() As String
    On Error Resume Next
    EngineGetTermFormatPref = Application.Run("PleadingsEngine.GetTermFormatPref")
    If Err.Number <> 0 Then
        Debug.Print "EngineGetTermFormatPref: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetTermFormatPref = "BOLD"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

# FILE: Rules_TextScan.bas

```vb
Attribute VB_Name = "Rules_TextScan"
' ============================================================
' Rules_TextScan.bas
' Combined text-scanning proofreading rules:
'   - Check_RepeatedWords (from Rule02)
'   - Check_SpellOutUnderTen (from Rule34)
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME_REPEATED As String = "repeated_words"
Private Const RULE_NAME_SPELL_OUT As String = "spell_out_under_ten"

' ============================================================
'  PUBLIC: Check_RepeatedWords
'  Detects consecutive repeated words (e.g. "the the").
'  Known-valid repetitions (e.g. "that that", "had had") are
'  flagged as "possible_error" rather than "error".
' ============================================================
Public Function Check_RepeatedWords(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraText As String
    Dim words() As String
    Dim cleanWords() As String
    Dim wordCount As Long
    Dim i As Long
    Dim prevWord As String
    Dim currWord As String
    Dim severity As String
    Dim issueText As String
    Dim suggestion As String
    Dim locStr As String
    Dim charPos As Long
    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim finding As Object
    Dim paraRange As Range

    ' -- Known-valid repetitions that may be intentional ---
    ' These get flagged as "possible_error" with a note
    ' to review context, rather than a hard "error".
    Dim knownValid As Variant
    knownValid = Array("that", "had", "is", "was", "can")

    ' -- Iterate all paragraphs ----------------------------
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph_RW
        End If

        ' Skip paragraphs outside the configured page range
        If Not EngineIsInPageRange(paraRange) Then
            GoTo NextParagraph_RW
        End If

        paraText = paraRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph_RW
        End If

        ' Skip very short or empty paragraphs
        If Len(Trim(paraText)) < 3 Then
            GoTo NextParagraph_RW
        End If

        ' Calculate auto-number prefix offset
        Dim rwListPrefixLen As Long
        rwListPrefixLen = GetSOListPrefixLen(para, paraText)

        ' -- Tokenise by scanning character positions directly ---
        ' This avoids misalignment from tabs, multiple spaces, NBSP.
        Dim tLen As Long
        tLen = Len(paraText)
        If tLen < 3 Then GoTo NextParagraph_RW

        prevWord = ""
        Dim prevTokenStart As Long, prevTokenEnd As Long
        prevTokenStart = 0: prevTokenEnd = 0

        Dim scanPos As Long
        scanPos = 1  ' 1-based position in paraText

        Do While scanPos <= tLen
            ' Skip whitespace
            Dim sc As String
            sc = Mid$(paraText, scanPos, 1)
            If sc = " " Or sc = vbTab Or sc = ChrW(160) Or _
               sc = vbCr Or sc = vbLf Or sc = Chr(11) Then
                scanPos = scanPos + 1
                GoTo NextScanPos_RW
            End If

            ' Found start of a token
            Dim tokStart As Long
            tokStart = scanPos
            Do While scanPos <= tLen
                sc = Mid$(paraText, scanPos, 1)
                If sc = " " Or sc = vbTab Or sc = ChrW(160) Or _
                   sc = vbCr Or sc = vbLf Or sc = Chr(11) Then Exit Do
                scanPos = scanPos + 1
            Loop
            Dim tokEnd As Long
            tokEnd = scanPos  ' one past end (exclusive)

            Dim rawToken As String
            rawToken = Mid$(paraText, tokStart, tokEnd - tokStart)
            currWord = LCase(StripPunctuation(rawToken))

            If Len(currWord) = 0 Then
                prevWord = ""
                GoTo NextScanPos_RW
            End If

            ' Check for repetition with previous token
            If currWord = prevWord And Len(currWord) > 0 Then
                ' Determine severity
                If IsKnownValidRepetition(currWord, knownValid) Then
                    severity = "possible_error"
                    issueText = "Repeated word '" & currWord & "' " & _
                                "-- review context; may be intentional"
                Else
                    severity = "error"
                    issueText = "Repeated word '" & currWord & "' detected"
                End If

                suggestion = "Remove the duplicate '" & currWord & "'"

                ' tokStart is 1-based in paraText; convert to document position
                rangeStart = paraRange.Start + (tokStart - 1) - rwListPrefixLen
                rangeEnd = rangeStart + (tokEnd - tokStart)

                Err.Clear
                Dim matchRange As Range
                Set matchRange = doc.Range(rangeStart, rangeEnd)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                Else
                    locStr = EngineGetLocationString(matchRange, doc)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    End If
                End If

                Set finding = CreateIssueDict(RULE_NAME_REPEATED, locStr, issueText, suggestion, rangeStart, rangeEnd, severity)
                issues.Add finding
            End If

            prevWord = currWord
            prevTokenStart = tokStart
            prevTokenEnd = tokEnd
NextScanPos_RW:
        Loop

NextParagraph_RW:
    Next para
    On Error GoTo 0

    Set Check_RepeatedWords = issues
End Function

' ============================================================
'  PUBLIC: Check_SpellOutUnderTen
'  In running prose, numbers under 10 should be written in
'  words (e.g. "seven" instead of "7").
' ============================================================
Public Function Check_SpellOutUnderTen(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraText As String
    Dim styleName As String
    Dim i As Long
    Dim ch As String
    Dim digitVal As Long
    Dim finding As Object
    Dim locStr As String
    Dim charRange As Range
    Dim textLen As Long

    ' Number word map
    Dim numberWords(0 To 9) As String
    numberWords(0) = "zero"
    numberWords(1) = "one"
    numberWords(2) = "two"
    numberWords(3) = "three"
    numberWords(4) = "four"
    numberWords(5) = "five"
    numberWords(6) = "six"
    numberWords(7) = "seven"
    numberWords(8) = "eight"
    numberWords(9) = "nine"

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph_SO
        End If

        ' Skip paragraphs outside the configured page range
        If Not EngineIsInPageRange(paraRange) Then
            GoTo NextParagraph_SO
        End If

        ' -- Check paragraph style for exclusions ------------
        styleName = ""
        styleName = paraRange.ParagraphStyle
        If Err.Number <> 0 Then
            Err.Clear
            styleName = ""
        End If

        If IsExcludedStyle(styleName) Then
            GoTo NextParagraph_SO
        End If

        ' -- Skip block quotes / indented extracts ----------
        Dim isBlockQ As Boolean
        isBlockQ = False
        isBlockQ = Application.Run("Rules_Formatting.IsBlockQuotePara", para)
        If Err.Number <> 0 Then isBlockQ = False: Err.Clear
        If isBlockQ Then GoTo NextParagraph_SO

        ' -- Get paragraph text ------------------------------
        paraText = paraRange.Text
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextParagraph_SO
        End If

        textLen = Len(paraText)
        If textLen = 0 Then GoTo NextParagraph_SO

        ' -- Calculate auto-number prefix offset -------------
        Dim soListPrefixLen As Long
        soListPrefixLen = GetSOListPrefixLen(para, paraText)

        ' -- Scan character by character for digits 0-9 ------
        For i = 1 To textLen
            ch = Mid(paraText, i, 1)

            ' Check if character is a digit 0-9
            If ch >= "0" And ch <= "9" Then
                digitVal = CInt(ch)

                ' -- Check: isolated digit (not part of larger number) --
                If IsPartOfLargerNumber(paraText, i, textLen) Then
                    GoTo NextChar
                End If

                ' -- Check: digit adjacent to a letter (postcodes, codes) --
                If IsAdjacentToLetter(paraText, i, textLen) Then
                    GoTo NextChar
                End If

                ' -- Check: preceded by structural reference word --
                If IsPrecededByStructuralRef(paraText, i) Then
                    GoTo NextChar
                End If

                ' -- Check: inside parentheses (clause sub-numbers) --
                If IsInsideParentheses(paraText, i) Then
                    GoTo NextChar
                End If

                ' -- Check: digit followed by opening bracket (clause ref like 1(4)) --
                If IsFollowedByBracket(paraText, i, textLen) Then
                    GoTo NextChar
                End If

                ' -- Check: digit followed by month name (date like 1 October) --
                If IsFollowedByMonthName(paraText, i, textLen) Then
                    GoTo NextChar
                End If

                ' -- Check: part of a range pattern --
                If IsPartOfRange(paraText, i, textLen) Then
                    GoTo NextChar
                End If

                ' -- Check: citation context --
                If IsInCitationContext(paraText, i) Then
                    GoTo NextChar
                End If

                ' -- Check: preceded by currency/unit symbols --
                If IsPrecededByCurrencyOrUnit(paraText, i) Then
                    GoTo NextChar
                End If

                ' -- Check: conjunction-linked structural ref --
                ' e.g. "paragraphs 4 and 5" — the "5" is preceded by "and"
                ' but the "4" before it has a structural ref
                If IsConjunctionLinkedRef(paraText, i) Then
                    GoTo NextChar
                End If

                ' -- All checks passed: flag this digit ------
                Dim rangeStart As Long
                Dim rangeEnd As Long

                rangeStart = paraRange.Start + i - 1 - soListPrefixLen
                rangeEnd = rangeStart + 1

                Err.Clear
                Set charRange = doc.Range(rangeStart, rangeEnd)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                Else
                    locStr = EngineGetLocationString(charRange, doc)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    End If
                End If

                Set finding = CreateIssueDict(RULE_NAME_SPELL_OUT, locStr, "Number under 10 is given as a figure in running prose.", "Write '" & numberWords(digitVal) & "' instead of '" & ch & "'.", rangeStart, rangeEnd, "warning", False)
                issues.Add finding
            End If

NextChar:
        Next i

NextParagraph_SO:
    Next para
    On Error GoTo 0

    Set Check_SpellOutUnderTen = issues
End Function

' ============================================================
'  HELPERS FOR Check_RepeatedWords
' ============================================================

' ------------------------------------------------------------
'  PRIVATE: Strip leading and trailing punctuation from a word
'  Removes characters like . , ; : ! ? " ' ( ) [ ] etc.
' ------------------------------------------------------------
Private Function StripPunctuation(ByVal word As String) As String
    Dim ch As String
    Dim startPos As Long
    Dim endPos As Long

    word = Trim(word)
    If Len(word) = 0 Then
        StripPunctuation = ""
        Exit Function
    End If

    ' Strip from start
    startPos = 1
    Do While startPos <= Len(word)
        ch = Mid(word, startPos, 1)
        If IsPunctuation(ch) Then
            startPos = startPos + 1
        Else
            Exit Do
        End If
    Loop

    ' Strip from end
    endPos = Len(word)
    Do While endPos >= startPos
        ch = Mid(word, endPos, 1)
        If IsPunctuation(ch) Then
            endPos = endPos - 1
        Else
            Exit Do
        End If
    Loop

    If startPos > endPos Then
        StripPunctuation = ""
    Else
        StripPunctuation = Mid(word, startPos, endPos - startPos + 1)
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if a character is punctuation
' ------------------------------------------------------------
Private Function IsPunctuation(ByVal ch As String) As Boolean
    Dim PUNCT_CHARS As String
    PUNCT_CHARS = ".,;:!?""'()[]{}/-" & Chr(8220) & Chr(8221) & _
                  Chr(8216) & Chr(8217) & Chr(8212) & Chr(8211)
    IsPunctuation = (InStr(1, PUNCT_CHARS, ch) > 0)
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if a word is in the known-valid list
' ------------------------------------------------------------
Private Function IsKnownValidRepetition(ByVal word As String, _
                                         ByRef knownValid As Variant) As Boolean
    Dim i As Long
    Dim lWord As String
    lWord = LCase(word)

    For i = LBound(knownValid) To UBound(knownValid)
        If LCase(CStr(knownValid(i))) = lWord Then
            IsKnownValidRepetition = True
            Exit Function
        End If
    Next i

    IsKnownValidRepetition = False
End Function

' ============================================================
'  HELPERS FOR Check_SpellOutUnderTen
' ============================================================

' ------------------------------------------------------------
'  PRIVATE: Check if paragraph style should be excluded
'  Excludes: Table, Code, Data, Technical, Footnote
' ------------------------------------------------------------
' Calculate the length of auto-generated list numbering text
' that appears in Range.Text but doesn't map to document positions.
Private Function GetSOListPrefixLen(para As Paragraph, ByVal paraText As String) As Long
    GetSOListPrefixLen = 0
    On Error Resume Next
    Dim lStr As String
    lStr = para.Range.ListFormat.ListString
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    On Error GoTo 0
    If Len(lStr) = 0 Then Exit Function
    If Len(paraText) > Len(lStr) Then
        If Left$(paraText, Len(lStr)) = lStr Then
            GetSOListPrefixLen = Len(lStr)
            If Mid$(paraText, GetSOListPrefixLen + 1, 1) = vbTab Then
                GetSOListPrefixLen = GetSOListPrefixLen + 1
            End If
        End If
    End If
End Function

Private Function IsExcludedStyle(ByVal styleName As String) As Boolean
    Dim lStyle As String
    lStyle = LCase(styleName)

    IsExcludedStyle = (InStr(lStyle, "table") > 0) Or _
                      (InStr(lStyle, "code") > 0) Or _
                      (InStr(lStyle, "data") > 0) Or _
                      (InStr(lStyle, "technical") > 0) Or _
                      (InStr(lStyle, "footnote") > 0)
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if the digit is part of a larger number
'  (preceded or followed by another digit or decimal point)
' ------------------------------------------------------------
Private Function IsPartOfLargerNumber(ByRef txt As String, _
                                       ByVal pos As Long, _
                                       ByVal textLen As Long) As Boolean
    Dim prevChar As String
    Dim nextChar As String

    IsPartOfLargerNumber = False

    ' Check character before
    If pos > 1 Then
        prevChar = Mid(txt, pos - 1, 1)
        If (prevChar >= "0" And prevChar <= "9") Or _
           prevChar = "." Or prevChar = "," Then
            IsPartOfLargerNumber = True
            Exit Function
        End If
    End If

    ' Check character after
    If pos < textLen Then
        nextChar = Mid(txt, pos + 1, 1)
        If (nextChar >= "0" And nextChar <= "9") Or _
           nextChar = "." Or nextChar = "," Then
            IsPartOfLargerNumber = True
            Exit Function
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is preceded by a structural
'  reference word (section, para, clause, etc.)
' ------------------------------------------------------------
Private Function IsPrecededByStructuralRef(ByRef txt As String, _
                                            ByVal pos As Long) As Boolean
    Dim refWords As Variant
    refWords = Array("section", "sect", "para", "paragraph", "clause", _
                     "article", "art", "rule", "reg", "regulation", _
                     "chapter", "page", "part", "schedule", "sch", _
                     "annex", "appendix", "item", "figure", "fig", _
                     "table", "tab", "footnote", "endnote", "version", _
                     "vol", "no", "ch", "cl", "fn", "pt", "pp", "p", "r", "s")

    IsPrecededByStructuralRef = False

    ' Extract the word immediately before the digit
    Dim prevWord As String
    prevWord = GetPrecedingWord(txt, pos)
    If Len(prevWord) = 0 Then Exit Function

    Dim lWord As String
    lWord = LCase(prevWord)

    ' Strip trailing "s" to handle plurals (e.g. "Rules" -> "rule")
    Dim lWordBase As String
    lWordBase = lWord
    If Len(lWordBase) > 2 And Right$(lWordBase, 1) = "s" Then
        lWordBase = Left$(lWordBase, Len(lWordBase) - 1)
    End If

    Dim j As Long
    For j = LBound(refWords) To UBound(refWords)
        If lWord = LCase(CStr(refWords(j))) Or _
           lWordBase = LCase(CStr(refWords(j))) Then
            IsPrecededByStructuralRef = True
            Exit Function
        End If
    Next j
End Function

' ------------------------------------------------------------
'  PRIVATE: Get the word immediately preceding position pos
'  Looks back from pos, skipping whitespace, then collecting
'  letters until a non-letter is found.
' ------------------------------------------------------------
Private Function GetPrecedingWord(ByRef txt As String, _
                                   ByVal pos As Long) As String
    Dim k As Long
    Dim ch As String
    Dim wordEnd As Long
    Dim wordStart As Long

    GetPrecedingWord = ""

    ' Skip whitespace before the digit
    k = pos - 1
    Do While k >= 1
        ch = Mid(txt, k, 1)
        If ch <> " " And ch <> vbTab Then Exit Do
        k = k - 1
    Loop

    If k < 1 Then Exit Function

    ' Check we landed on a letter or full stop (for abbreviations like "s.")
    ' Skip trailing full stop/dot
    If ch = "." Then
        k = k - 1
        If k < 1 Then Exit Function
    End If

    ' Now collect the word (letters only) going backwards
    wordEnd = k
    Do While k >= 1
        ch = Mid(txt, k, 1)
        If IsLetterChar(ch) Then
            k = k - 1
        Else
            Exit Do
        End If
    Loop
    wordStart = k + 1

    If wordStart > wordEnd Then Exit Function

    GetPrecedingWord = Mid(txt, wordStart, wordEnd - wordStart + 1)
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is inside parentheses -- catches
'  clause sub-numbers like "34(3)(e)", "(iv)", "s.2(1)" etc.
' ------------------------------------------------------------
Private Function IsInsideParentheses(ByRef txt As String, _
                                      ByVal pos As Long) As Boolean
    IsInsideParentheses = False

    ' Check for opening paren before (skipping digits and letters)
    Dim k As Long
    k = pos - 1
    If k >= 1 Then
        If Mid(txt, k, 1) = "(" Then
            IsInsideParentheses = True
            Exit Function
        End If
    End If

    ' Check for closing paren after (skipping ahead past the digit)
    k = pos + 1
    If k <= Len(txt) Then
        If Mid(txt, k, 1) = ")" Then
            IsInsideParentheses = True
            Exit Function
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is part of a range pattern
'  e.g. "7-12", "3--9", digit followed by en-dash/hyphen
'  and another digit, or preceded by digit+dash
' ------------------------------------------------------------
Private Function IsPartOfRange(ByRef txt As String, _
                                ByVal pos As Long, _
                                ByVal textLen As Long) As Boolean
    Dim nextPos As Long
    Dim nextChar As String
    Dim prevPos As Long
    Dim prevChar As String

    IsPartOfRange = False

    ' Check forward: digit followed by dash/en-dash then digit
    nextPos = pos + 1
    If nextPos <= textLen Then
        nextChar = Mid(txt, nextPos, 1)
        ' Hyphen, en-dash (ChrW(8211)), or em-dash (ChrW(8212))
        If nextChar = "-" Or AscW(nextChar) = 8211 Or AscW(nextChar) = 8212 Then
            ' Check if next-next is a digit
            If nextPos + 1 <= textLen Then
                Dim afterDash As String
                afterDash = Mid(txt, nextPos + 1, 1)
                If afterDash >= "0" And afterDash <= "9" Then
                    IsPartOfRange = True
                    Exit Function
                End If
            End If
        End If
    End If

    ' Check backward: preceded by dash then digit (we are the end of a range)
    prevPos = pos - 1
    If prevPos >= 1 Then
        prevChar = Mid(txt, prevPos, 1)
        If prevChar = "-" Or AscW(prevChar) = 8211 Or AscW(prevChar) = 8212 Then
            If prevPos - 1 >= 1 Then
                Dim beforeDash As String
                beforeDash = Mid(txt, prevPos - 1, 1)
                If beforeDash >= "0" And beforeDash <= "9" Then
                    IsPartOfRange = True
                    Exit Function
                End If
            End If
        End If
    End If

    ' Check for "to" pattern: digit + space + "to" + space + digit
    ' Forward check -- need at least 5 chars after pos: " to X"
    If pos + 5 <= textLen Then
        If Mid(txt, pos + 1, 4) = " to " Then
            Dim afterTo As String
            afterTo = Mid(txt, pos + 5, 1)
            If afterTo >= "0" And afterTo <= "9" Then
                IsPartOfRange = True
                Exit Function
            End If
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is in a citation context
'  Look for "[" within 10 characters before
' ------------------------------------------------------------
Private Function IsInCitationContext(ByRef txt As String, _
                                      ByVal pos As Long) As Boolean
    Dim startSearch As Long
    Dim k As Long

    IsInCitationContext = False

    startSearch = pos - 10
    If startSearch < 1 Then startSearch = 1

    For k = startSearch To pos - 1
        If Mid(txt, k, 1) = "[" Then
            IsInCitationContext = True
            Exit Function
        End If
    Next k
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is preceded by currency symbols,
'  percentage, or unit markers
' ------------------------------------------------------------
Private Function IsPrecededByCurrencyOrUnit(ByRef txt As String, _
                                             ByVal pos As Long) As Boolean
    Dim prevChar As String
    Dim prevCode As Long

    IsPrecededByCurrencyOrUnit = False

    If pos <= 1 Then Exit Function

    prevChar = Mid(txt, pos - 1, 1)
    prevCode = AscW(prevChar)

    ' Currency symbols: $, pound sign (163), euro (8364), yen (165)
    ' Unit markers: %, #
    Select Case prevCode
        Case 36    ' $
            IsPrecededByCurrencyOrUnit = True
        Case 163   ' pound sign
            IsPrecededByCurrencyOrUnit = True
        Case 8364  ' euro sign
            IsPrecededByCurrencyOrUnit = True
        Case 165   ' yen sign
            IsPrecededByCurrencyOrUnit = True
        Case 37    ' %
            IsPrecededByCurrencyOrUnit = True
        Case 35    ' #
            IsPrecededByCurrencyOrUnit = True
    End Select

    ' Also check if the character after the digit is %
    If Not IsPrecededByCurrencyOrUnit Then
        If pos < Len(txt) Then
            Dim nextChar As String
            nextChar = Mid(txt, pos + 1, 1)
            If nextChar = "%" Then
                IsPrecededByCurrencyOrUnit = True
            End If
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is linked via conjunction (and/or/to)
'  to another digit that IS preceded by a structural reference.
'  Catches patterns like "paragraphs 4 and 5", "rules 3 to 7",
'  "sections 2 or 3", "paragraphs 4, 5 and 6".
' ------------------------------------------------------------
Private Function IsConjunctionLinkedRef(ByRef txt As String, _
                                         ByVal pos As Long) As Boolean
    IsConjunctionLinkedRef = False

    ' Get the word before this digit
    Dim prevWord As String
    prevWord = LCase(GetPrecedingWord(txt, pos))
    If Len(prevWord) = 0 Then Exit Function

    ' Must be preceded by "and", "or", "to", or a comma
    Dim isConj As Boolean
    isConj = (prevWord = "and" Or prevWord = "or" Or prevWord = "to")

    ' Also handle comma-separated: "paragraphs 4, 5 and 6"
    If Not isConj Then
        ' Check if preceded by comma (skip spaces)
        Dim k As Long
        k = pos - 1
        Do While k >= 1
            Dim c As String
            c = Mid$(txt, k, 1)
            If c = " " Or c = vbTab Then
                k = k - 1
            Else
                Exit Do
            End If
        Loop
        If k >= 1 And Mid$(txt, k, 1) = "," Then
            isConj = True
        End If
    End If

    If Not isConj Then Exit Function

    ' Now scan backwards past the conjunction to find a preceding digit
    ' For "and"/"or"/"to": skip back past the conjunction word + spaces + the digit
    ' For comma: already at the comma, skip back past it + spaces + the digit
    Dim scanPos As Long
    scanPos = pos

    ' Skip back to before the preceding word / comma
    scanPos = scanPos - 1  ' space before digit
    Do While scanPos >= 1 And (Mid$(txt, scanPos, 1) = " " Or Mid$(txt, scanPos, 1) = vbTab)
        scanPos = scanPos - 1
    Loop
    ' Skip back past the conjunction word or comma
    If isConj And (prevWord = "and" Or prevWord = "or" Or prevWord = "to") Then
        scanPos = scanPos - Len(prevWord)
    ElseIf isConj Then
        ' comma case — scanPos is already past the comma
    End If
    ' Skip spaces before the conjunction
    Do While scanPos >= 1 And (Mid$(txt, scanPos, 1) = " " Or Mid$(txt, scanPos, 1) = vbTab)
        scanPos = scanPos - 1
    Loop

    ' Check if there's a digit at scanPos
    If scanPos >= 1 Then
        Dim prevCh As String
        prevCh = Mid$(txt, scanPos, 1)
        If prevCh >= "0" And prevCh <= "9" Then
            ' Found a digit — check if THAT digit is preceded by a structural ref
            If IsPrecededByStructuralRef(txt, scanPos) Then
                IsConjunctionLinkedRef = True
                Exit Function
            End If
            ' Or if THAT digit is also conjunction-linked (recursive chain)
            If IsConjunctionLinkedRef(txt, scanPos) Then
                IsConjunctionLinkedRef = True
            End If
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is adjacent to a letter
'  (postcodes like SO50 2ZH, codes like ET1, etc.)
' ------------------------------------------------------------
Private Function IsAdjacentToLetter(ByRef txt As String, _
                                     ByVal pos As Long, _
                                     ByVal textLen As Long) As Boolean
    IsAdjacentToLetter = False

    ' Check character before
    If pos > 1 Then
        If IsLetterChar(Mid(txt, pos - 1, 1)) Then
            IsAdjacentToLetter = True
            Exit Function
        End If
    End If

    ' Check character after
    If pos < textLen Then
        If IsLetterChar(Mid(txt, pos + 1, 1)) Then
            IsAdjacentToLetter = True
            Exit Function
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is followed by opening bracket
'  (clause references like 1(4), 3(a), etc.)
' ------------------------------------------------------------
Private Function IsFollowedByBracket(ByRef txt As String, _
                                      ByVal pos As Long, _
                                      ByVal textLen As Long) As Boolean
    IsFollowedByBracket = False

    If pos < textLen Then
        If Mid(txt, pos + 1, 1) = "(" Then
            IsFollowedByBracket = True
        End If
    End If
End Function

' ------------------------------------------------------------
'  PRIVATE: Check if digit is followed by a month name
'  (date patterns like "1 October 2004")
' ------------------------------------------------------------
Private Function IsFollowedByMonthName(ByRef txt As String, _
                                        ByVal pos As Long, _
                                        ByVal textLen As Long) As Boolean
    IsFollowedByMonthName = False

    ' Need at least a space + 3 chars after the digit
    If pos + 4 > textLen Then Exit Function

    ' Must be followed by a space
    If Mid(txt, pos + 1, 1) <> " " Then Exit Function

    ' Extract the next word after the space
    Dim wordStart As Long
    wordStart = pos + 2
    Dim wordEnd As Long
    wordEnd = wordStart
    Do While wordEnd <= textLen
        If Not IsLetterChar(Mid(txt, wordEnd, 1)) Then Exit Do
        wordEnd = wordEnd + 1
    Loop

    If wordEnd <= wordStart Then Exit Function

    Dim nextWord As String
    nextWord = LCase(Mid(txt, wordStart, wordEnd - wordStart))

    Dim months As Variant
    months = Array("january", "february", "march", "april", "may", _
                   "june", "july", "august", "september", "october", _
                   "november", "december")

    Dim m As Long
    For m = LBound(months) To UBound(months)
        If nextWord = CStr(months(m)) Then
            IsFollowedByMonthName = True
            Exit Function
        End If
    Next m
End Function

' ============================================================
'  SHARED HELPER (used by both rules' helpers)
' ============================================================

' ------------------------------------------------------------
'  PRIVATE: Check if a character is a letter (A-Z, a-z,
'  extended Latin)
' ------------------------------------------------------------
Private Function IsLetterChar(ByVal ch As String) As Boolean
    Dim code As Long
    code = AscW(ch)
    IsLetterChar = (code >= 65 And code <= 90) Or _
                   (code >= 97 And code <= 122) Or _
                   (code >= 192 And code <= 687) ' Extended Latin
End Function


' ----------------------------------------------------------------
'  PRIVATE: Create a dictionary-based finding (no class dependency)
' ----------------------------------------------------------------
Private Function CreateIssueDict(ByVal ruleName_ As String, _
                                 ByVal location_ As String, _
                                 ByVal issue_ As String, _
                                 ByVal suggestion_ As String, _
                                 ByVal rangeStart_ As Long, _
                                 ByVal rangeEnd_ As Long, _
                                 Optional ByVal severity_ As String = "error", _
                                 Optional ByVal autoFixSafe_ As Boolean = False) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("RuleName") = ruleName_
    d("Location") = location_
    d("Issue") = issue_
    d("Suggestion") = suggestion_
    d("RangeStart") = rangeStart_
    d("RangeEnd") = rangeEnd_
    d("Severity") = severity_
    d("AutoFixSafe") = autoFixSafe_
    Set CreateIssueDict = d
End Function


' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineIsInPageRange = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.GetLocationString
' ----------------------------------------------------------------
Private Function EngineGetLocationString(rng As Object, doc As Document) As String
    On Error Resume Next
    EngineGetLocationString = Application.Run("PleadingsEngine.GetLocationString", rng, doc)
    If Err.Number <> 0 Then
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function

```

