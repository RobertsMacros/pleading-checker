VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPleadingsChecker
   Caption         =   "Pleadings Checker"
   ClientHeight    =   500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   700
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
Private WithEvents btnEditBrand     As MSForms.CommandButton
Private WithEvents btnSaveBrands    As MSForms.CommandButton
Private WithEvents btnLoadBrands    As MSForms.CommandButton

Private fraRules        As MSForms.Frame
Private WithEvents txtPageRange As MSForms.TextBox
Private lstBrands       As MSForms.ListBox
Private txtBrandCorrect As MSForms.TextBox
Private WithEvents txtBrandIncorrect As MSForms.TextBox
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
Private targetDoc       As Document
Private editingBrandIndex As Long      ' -1 = not editing; >= 0 = list index being edited
Private placeholderActive As Boolean   ' True if brand placeholder text is showing
Private pageRangePlaceholderActive As Boolean  ' True if page range placeholder is showing

' ============================================================
'  FORM INITIALISATION -- creates all controls at runtime
' ============================================================
Private Sub UserForm_Initialize()
    editingBrandIndex = -1
    placeholderActive = False
    pageRangePlaceholderActive = False

    Dim lbl As MSForms.Label
    Dim yPos As Single

    ' -- Overall form padding ----------------------------------
    Const PAD As Single = 10
    Const FULL_W As Single = 680     ' narrower, more compact form
    Const BTN_W As Single = 78
    Const BTN_H As Single = 22
    Const TXT_H As Single = 20
    Const CHK_H As Single = 16
    Const LBL_H As Single = 14
    Const SEC_GAP As Single = 6      ' gap between sections
    Const ITEM_GAP As Single = 2     ' gap within sections

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
        .Left = PAD: .Top = yPos: .Width = 40: .Height = LBL_H
        .Font.Size = 9: .Font.Bold = True
    End With

    Set btnSelectAll = Me.Controls.Add("Forms.CommandButton.1", "btnSelectAll")
    With btnSelectAll
        .Caption = "Select All"
        .Left = PAD + 44: .Top = yPos - 1: .Width = 62: .Height = 18
        .Font.Size = 7
    End With

    Set btnDeselectAll = Me.Controls.Add("Forms.CommandButton.1", "btnDeselectAll")
    With btnDeselectAll
        .Caption = "Deselect All"
        .Left = PAD + 44 + 64: .Top = yPos - 1: .Width = 62: .Height = 18
        .Font.Size = 7
    End With

    yPos = yPos + 18 + ITEM_GAP

    ' ==========================================================
    '  ROW 2: Rule checkboxes in scrollable frame
    ' ==========================================================
    Set fraRules = Me.Controls.Add("Forms.Frame.1", "fraRules")
    With fraRules
        .Caption = ""
        .Left = PAD: .Top = yPos
        .Width = FULL_W
        .Height = 80
        .ScrollBars = fmScrollBarsVertical
        .KeepScrollBarsVisible = fmScrollBarsVertical
    End With

    BuildRuleCheckboxList nRules

    yPos = yPos + fraRules.Height + SEC_GAP

    ' ==========================================================
    '  ROW 3: Left column (Page Range + Brand Rules)
    '         Right column (Options)
    ' ==========================================================
    Dim colLeft As Single
    Dim colRight As Single
    Dim leftW As Single
    colLeft = PAD
    leftW = FULL_W * 0.52
    colRight = PAD + leftW + SEC_GAP
    Dim row3Top As Single
    row3Top = yPos

    ' ---- LEFT COLUMN: Page Range ----
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblPageHeader")
    With lbl
        .Caption = "Page Range"
        .Left = colLeft: .Top = yPos: .Width = 120: .Height = LBL_H
        .Font.Size = 9: .Font.Bold = True
    End With
    yPos = yPos + LBL_H + ITEM_GAP

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblPageRange")
    With lbl
        .Caption = "Pages:"
        .Left = colLeft: .Top = yPos + 2: .Width = 36: .Height = LBL_H
    End With

    Set txtPageRange = Me.Controls.Add("Forms.TextBox.1", "txtPageRange")
    With txtPageRange
        .Left = colLeft + 36: .Top = yPos: .Width = leftW - 40: .Height = TXT_H
        .Text = ""
    End With
    ShowPageRangePlaceholder

    yPos = yPos + TXT_H + SEC_GAP

    ' ---- LEFT COLUMN: Brand Rules ----
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblBrandHeader")
    With lbl
        .Caption = "Brand Rules"
        .Left = colLeft: .Top = yPos: .Width = 120: .Height = LBL_H
        .Font.Size = 9: .Font.Bold = True
    End With
    yPos = yPos + LBL_H + ITEM_GAP

    Dim brandListW As Single
    brandListW = leftW - BTN_W - ITEM_GAP - 4

    Set lstBrands = Me.Controls.Add("Forms.ListBox.1", "lstBrands")
    With lstBrands
        .Left = colLeft: .Top = yPos
        .Width = brandListW: .Height = 56
        .Font.Size = 7.5
    End With

    ' Brand action buttons (right of list, stacked)
    Dim btnX As Single
    btnX = colLeft + brandListW + ITEM_GAP
    Dim brandBtnY As Single
    brandBtnY = yPos

    Set btnAddBrand = Me.Controls.Add("Forms.CommandButton.1", "btnAddBrand")
    With btnAddBrand
        .Caption = "Add"
        .Left = btnX: .Top = brandBtnY: .Width = BTN_W: .Height = BTN_H
        .Font.Size = 7.5
    End With
    brandBtnY = brandBtnY + BTN_H + 1

    Set btnRemoveBrand = Me.Controls.Add("Forms.CommandButton.1", "btnRemoveBrand")
    With btnRemoveBrand
        .Caption = "Remove"
        .Left = btnX: .Top = brandBtnY: .Width = BTN_W / 2 - 1: .Height = BTN_H
        .Font.Size = 7
    End With

    Set btnEditBrand = Me.Controls.Add("Forms.CommandButton.1", "btnEditBrand")
    With btnEditBrand
        .Caption = "Edit"
        .Left = btnX + BTN_W / 2 + 1: .Top = brandBtnY: .Width = BTN_W / 2 - 1: .Height = BTN_H
        .Font.Size = 7
    End With

    yPos = yPos + lstBrands.Height + ITEM_GAP

    ' Brand input row: Correct + Incorrect + Save/Load
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblCorrectForm")
    With lbl
        .Caption = "Correct:"
        .Left = colLeft: .Top = yPos + 2: .Width = 42: .Height = LBL_H
        .Font.Size = 7.5
    End With

    Set txtBrandCorrect = Me.Controls.Add("Forms.TextBox.1", "txtBrandCorrect")
    With txtBrandCorrect
        .Left = colLeft + 42: .Top = yPos: .Width = 90: .Height = TXT_H
        .Font.Size = 7.5
    End With

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblIncorrectVars")
    With lbl
        .Caption = "Variants:"
        .Left = colLeft + 136: .Top = yPos + 2: .Width = 42: .Height = LBL_H
        .Font.Size = 7.5
    End With

    Set txtBrandIncorrect = Me.Controls.Add("Forms.TextBox.1", "txtBrandIncorrect")
    With txtBrandIncorrect
        .Left = colLeft + 178: .Top = yPos: .Width = 100: .Height = TXT_H
        .Font.Size = 7.5
    End With
    ShowBrandPlaceholder

    ' Save/Load buttons inline after inputs
    Dim slX As Single
    slX = colLeft + 282

    Set btnSaveBrands = Me.Controls.Add("Forms.CommandButton.1", "btnSaveBrands")
    With btnSaveBrands
        .Caption = "Save"
        .Left = slX: .Top = yPos: .Width = 36: .Height = TXT_H
        .Font.Size = 7
    End With

    Set btnLoadBrands = Me.Controls.Add("Forms.CommandButton.1", "btnLoadBrands")
    With btnLoadBrands
        .Caption = "Load"
        .Left = slX + 38: .Top = yPos: .Width = 36: .Height = TXT_H
        .Font.Size = 7
    End With

    Dim leftBottomY As Single
    leftBottomY = yPos + TXT_H

    ' ---- RIGHT COLUMN: Options (starting from row3Top) ----
    Dim optY As Single
    optY = row3Top

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblOptionsHeader")
    With lbl
        .Caption = "Options"
        .Left = colRight: .Top = optY: .Width = 120: .Height = LBL_H
        .Font.Size = 9: .Font.Bold = True
    End With
    optY = optY + LBL_H + ITEM_GAP

    Set chkAddComments = Me.Controls.Add("Forms.CheckBox.1", "chkAddComments")
    With chkAddComments
        .Caption = "Add comments"
        .Left = colRight: .Top = optY: .Width = 140: .Height = CHK_H
        .Value = True
        .Font.Size = 7.5
    End With
    optY = optY + CHK_H + ITEM_GAP

    Set chkTrackedChanges = Me.Controls.Add("Forms.CheckBox.1", "chkTrackedChanges")
    With chkTrackedChanges
        .Caption = "Tracked changes"
        .Left = colRight: .Top = optY: .Width = 140: .Height = CHK_H
        .Value = True
        .Font.Size = 7.5
    End With
    optY = optY + CHK_H + ITEM_GAP

    ' Spelling mode
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblSpellingMode")
    With lbl
        .Caption = "Spelling:"
        .Left = colRight: .Top = optY + 1: .Width = 52: .Height = LBL_H
        .Font.Size = 7.5
    End With

    Set optSpellingUK = Me.Controls.Add("Forms.OptionButton.1", "optSpellingUK")
    With optSpellingUK
        .Caption = "UK": .Left = colRight + 52: .Top = optY: .Width = 40: .Height = CHK_H
        .Value = True: .GroupName = "SpellingMode": .Font.Size = 7.5
    End With

    Set optSpellingUS = Me.Controls.Add("Forms.OptionButton.1", "optSpellingUS")
    With optSpellingUS
        .Caption = "US": .Left = colRight + 94: .Top = optY: .Width = 40: .Height = CHK_H
        .Value = False: .GroupName = "SpellingMode": .Font.Size = 7.5
    End With
    optY = optY + CHK_H + ITEM_GAP

    ' Date format
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblDateFormat")
    With lbl
        .Caption = "Date:"
        .Left = colRight: .Top = optY + 1: .Width = 52: .Height = LBL_H
        .Font.Size = 7.5
    End With

    Set optDateUK = Me.Controls.Add("Forms.OptionButton.1", "optDateUK")
    With optDateUK
        .Caption = "UK": .Left = colRight + 52: .Top = optY: .Width = 40: .Height = CHK_H
        .Value = True: .GroupName = "DateFormat": .Font.Size = 7.5
    End With

    Set optDateUS = Me.Controls.Add("Forms.OptionButton.1", "optDateUS")
    With optDateUS
        .Caption = "US": .Left = colRight + 94: .Top = optY: .Width = 40: .Height = CHK_H
        .Value = False: .GroupName = "DateFormat": .Font.Size = 7.5
    End With
    optY = optY + CHK_H + ITEM_GAP

    ' After full stop spacing
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblSpaceStyle")
    With lbl
        .Caption = "After full stop:"
        .Left = colRight: .Top = optY + 1: .Width = 72: .Height = LBL_H
        .Font.Size = 7.5
    End With

    Set cboSpaceStyle = Me.Controls.Add("Forms.ComboBox.1", "cboSpaceStyle")
    With cboSpaceStyle
        .Left = colRight + 72: .Top = optY: .Width = 80: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "One space"
        .AddItem "Two spaces"
        .ListIndex = 0
        .Font.Size = 7.5
    End With
    optY = optY + TXT_H + ITEM_GAP

    ' Outer quotes
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblQuoteNesting")
    With lbl
        .Caption = "Outer quotes:"
        .Left = colRight: .Top = optY + 1: .Width = 72: .Height = LBL_H
        .Font.Size = 7.5
    End With

    Set optQuoteSingle = Me.Controls.Add("Forms.OptionButton.1", "optQuoteSingle")
    With optQuoteSingle
        .Caption = "Single": .Left = colRight + 72: .Top = optY: .Width = 50: .Height = CHK_H
        .Value = True: .GroupName = "QuoteNesting": .Font.Size = 7.5
    End With

    Set optQuoteDouble = Me.Controls.Add("Forms.OptionButton.1", "optQuoteDouble")
    With optQuoteDouble
        .Caption = "Double": .Left = colRight + 124: .Top = optY: .Width = 50: .Height = CHK_H
        .Value = False: .GroupName = "QuoteNesting": .Font.Size = 7.5
    End With
    optY = optY + CHK_H + ITEM_GAP

    ' Smart quotes
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblSmartQuotes")
    With lbl
        .Caption = "Smart quotes:"
        .Left = colRight: .Top = optY + 1: .Width = 72: .Height = LBL_H
        .Font.Size = 7.5
    End With

    Set optSmart = Me.Controls.Add("Forms.OptionButton.1", "optSmart")
    With optSmart
        .Caption = "Smart": .Left = colRight + 72: .Top = optY: .Width = 50: .Height = CHK_H
        .Value = True: .GroupName = "SmartQuotes": .Font.Size = 7.5
    End With

    Set optSmartStraight = Me.Controls.Add("Forms.OptionButton.1", "optSmartStraight")
    With optSmartStraight
        .Caption = "Straight": .Left = colRight + 124: .Top = optY: .Width = 56: .Height = CHK_H
        .Value = False: .GroupName = "SmartQuotes": .Font.Size = 7.5
    End With
    optY = optY + CHK_H + ITEM_GAP

    ' Defined terms
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblDefinedTerms")
    With lbl
        .Caption = "Def. terms:"
        .Left = colRight: .Top = optY + 1: .Width = 56: .Height = LBL_H
        .Font.Size = 7.5
    End With

    Set cboTermFormat = Me.Controls.Add("Forms.ComboBox.1", "cboTermFormat")
    With cboTermFormat
        .Left = colRight + 56: .Top = optY: .Width = 70: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "Bold"
        .AddItem "Bold Italics"
        .AddItem "Italics"
        .AddItem "None"
        .ListIndex = 0
        .Font.Size = 7.5
    End With

    Dim lblAnd As MSForms.Label
    Set lblAnd = Me.Controls.Add("Forms.Label.1", "lblTermAnd")
    With lblAnd
        .Caption = "+"
        .Left = colRight + 128: .Top = optY + 1: .Width = 10: .Height = LBL_H
        .Font.Size = 7.5
    End With

    Set cboTermQuotes = Me.Controls.Add("Forms.ComboBox.1", "cboTermQuotes")
    With cboTermQuotes
        .Left = colRight + 140: .Top = optY: .Width = 80: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "Single quotes"
        .AddItem "Double quotes"
        .ListIndex = 1
        .Font.Size = 7.5
    End With

    ' Use the taller of left-column or right-column bottoms
    Dim row3BottomY As Single
    If leftBottomY > optY Then row3BottomY = leftBottomY Else row3BottomY = optY
    yPos = row3BottomY + SEC_GAP

    ' ==========================================================
    '  ROW 4: Action Buttons
    ' ==========================================================
    Const ACT_BTN_H As Single = 28
    Const ACT_BTN_W As Single = 100
    Const ACT_GAP As Single = 8

    Set btnRun = Me.Controls.Add("Forms.CommandButton.1", "btnRun")
    With btnRun
        .Caption = "Run Checks"
        .Left = PAD: .Top = yPos: .Width = ACT_BTN_W: .Height = ACT_BTN_H
        .Font.Bold = True
    End With

    Set btnExport = Me.Controls.Add("Forms.CommandButton.1", "btnExport")
    With btnExport
        .Caption = "Export Report"
        .Left = PAD + ACT_BTN_W + ACT_GAP: .Top = yPos
        .Width = ACT_BTN_W: .Height = ACT_BTN_H
    End With

    Set btnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    With btnClose
        .Caption = "Close"
        .Left = PAD + 2 * (ACT_BTN_W + ACT_GAP): .Top = yPos
        .Width = 70: .Height = ACT_BTN_H
    End With

    yPos = yPos + ACT_BTN_H + ITEM_GAP

    ' ==========================================================
    '  ROW 5: Status Bar
    ' ==========================================================
    Set lblStatus = Me.Controls.Add("Forms.Label.1", "lblStatus")
    With lblStatus
        .Caption = "Ready. Select rules and click Run."
        .Left = PAD: .Top = yPos: .Width = FULL_W: .Height = LBL_H
        .Font.Size = 8
    End With

    ' -- Load brand list ---------------------------------------
    RefreshBrandList

    ' -- Final form size based on layout ---
    Dim neededH As Single
    neededH = yPos + LBL_H + PAD
    If neededH < 300 Then neededH = 300

    Me.Width = FULL_W + 2 * PAD
    Me.Height = neededH

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

    ' Guard against InsideWidth returning zero or implausibly small values
    ' on some Word hosts during early initialisation
    Const MIN_USABLE_W As Single = 120   ' absolute floor (30 pts per column)
    Dim usableW As Single
    On Error Resume Next
    usableW = fraRules.InsideWidth
    If Err.Number <> 0 Then usableW = 0: Err.Clear
    On Error GoTo 0
    If usableW <= 0 Then usableW = fraRules.Width - COL_PAD * 2
    If usableW < MIN_USABLE_W Then usableW = MIN_USABLE_W

    Dim colW As Single
    colW = (usableW - COL_PAD * 2) / COLS

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
    Set targetDoc = PleadingsEngine.GetTargetDocument()
    If targetDoc Is Nothing Then
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

    ' Set page range from flexible input (ignore placeholder text)
    PleadingsEngine.SetPageRangeFromString GetPageRangeText()

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

    Set lastResults = PleadingsEngine.RunAllPleadingsRules(targetDoc, ruleConfig)

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

        Dim reply As VbMsgBoxResult
        reply = MsgBox(summary & errMsg & vbCrLf & vbCrLf & _
               "Apply suggestions to the document?", _
               vbYesNo + vbQuestion, "Pleadings Checker")

        If reply = vbYes Then
            lblStatus.Caption = "Applying suggestions..."
            Me.Repaint
            DoEvents

            Dim addComments As Boolean
            addComments = (chkAddComments.Value = True)

            If chkTrackedChanges.Value = True Then
                PleadingsEngine.ApplySuggestionsAsTrackedChanges targetDoc, lastResults, addComments
            Else
                PleadingsEngine.ApplyHighlights targetDoc, lastResults, addComments
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

    ' Fallback to temp directory if no valid path yet
    If Len(reportPath) = 0 Then
        reportPath = GetTempReportPath(sep)
    End If

    ' Ensure parent directory exists before writing
    Dim reportDir As String
    reportDir = GetParentDirectory(reportPath)
    If Len(reportDir) > 0 Then
        EnsureDirectoryExists reportDir
    End If

    lblStatus.Caption = "Exporting report..."
    Me.Repaint
    DoEvents

    Dim summary As String
    summary = PleadingsEngine.GenerateReport(lastResults, reportPath, targetDoc)

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

    ' Build informative export message
    Dim errCount As Long
    errCount = PleadingsEngine.GetRuleErrorCount()

    Dim msg As String
    msg = "Report saved to:" & vbCrLf & reportPath

    If logSaved And Len(logPath) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Debug log saved to:" & vbCrLf & logPath
    ElseIf DEBUG_MODE And Not logSaved Then
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
    GetTempReportPath = GetWritableTempDir() & sep & "pleadings_report.json"
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
    incorrectVariants = GetBrandIncorrectText()

    If correctForm = "" Or incorrectVariants = "" Then
        MsgBox "Enter both the correct form and at least one incorrect variant.", _
               vbExclamation, "Brand Rules"
        Exit Sub
    End If

    ' Normalise comma-separated variants: trim and remove blanks
    incorrectVariants = NormaliseBrandVariants(incorrectVariants)
    If Len(incorrectVariants) = 0 Then
        MsgBox "Enter at least one incorrect variant.", vbExclamation, "Brand Rules"
        Exit Sub
    End If

    If editingBrandIndex >= 0 Then
        ' Remove old rule first, then add updated
        Dim oldEntry As String
        oldEntry = lstBrands.List(editingBrandIndex)
        Dim oldCorrect As String
        oldCorrect = Left(oldEntry, InStr(oldEntry, " -> ") - 1)
        On Error Resume Next
        Application.Run "Rules_Brands.RemoveBrandRule", oldCorrect
        Err.Clear
        On Error GoTo 0
        editingBrandIndex = -1
        btnAddBrand.Caption = "Add"
    End If

    On Error Resume Next
    Application.Run "Rules_Brands.AddBrandRule", correctForm, incorrectVariants
    If Err.Number <> 0 Then
        MsgBox "Brand rules module not loaded.", vbExclamation, "Brand Rules"
        Err.Clear
    End If
    On Error GoTo 0

    txtBrandCorrect.Text = ""
    ClearBrandIncorrect
    ShowBrandPlaceholder
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

Private Sub btnEditBrand_Click()
    If lstBrands.ListIndex < 0 Then
        MsgBox "Please select a rule to edit.", vbInformation, "Brand Rules"
        Exit Sub
    End If

    Dim entry As String
    entry = lstBrands.List(lstBrands.ListIndex)
    Dim arrowPos As Long
    arrowPos = InStr(entry, " -> ")
    If arrowPos = 0 Then Exit Sub

    editingBrandIndex = lstBrands.ListIndex
    txtBrandCorrect.Text = Left(entry, arrowPos - 1)
    HideBrandPlaceholder
    txtBrandIncorrect.Text = Mid(entry, arrowPos + 4)
    btnAddBrand.Caption = "Save Edit"
End Sub

' -- Placeholder helpers for txtBrandIncorrect --
Private Sub ShowBrandPlaceholder()
    If txtBrandIncorrect Is Nothing Then Exit Sub
    If Len(Trim(txtBrandIncorrect.Text)) = 0 Or placeholderActive Then
        txtBrandIncorrect.Text = "e.g. colour, colur, coulour"
        txtBrandIncorrect.ForeColor = &HC0C0C0  ' light grey
        placeholderActive = True
    End If
End Sub

Private Sub HideBrandPlaceholder()
    If placeholderActive Then
        txtBrandIncorrect.Text = ""
        txtBrandIncorrect.ForeColor = &H0  ' black
        placeholderActive = False
    End If
End Sub

' Return the actual text, ignoring placeholder
Private Function GetBrandIncorrectText() As String
    If placeholderActive Then
        GetBrandIncorrectText = ""
    Else
        GetBrandIncorrectText = Trim(txtBrandIncorrect.Text)
    End If
End Function

Private Sub ClearBrandIncorrect()
    txtBrandIncorrect.Text = ""
    txtBrandIncorrect.ForeColor = &H0
    placeholderActive = False
End Sub

' Normalise comma-separated variants: trim each, remove blanks
Private Function NormaliseBrandVariants(ByVal raw As String) As String
    Dim parts() As String
    parts = Split(raw, ",")
    Dim result As String
    Dim p As Long
    For p = LBound(parts) To UBound(parts)
        Dim item As String
        item = Trim(parts(p))
        If Len(item) > 0 Then
            If Len(result) > 0 Then result = result & ", "
            result = result & item
        End If
    Next p
    NormaliseBrandVariants = result
End Function

Private Sub txtBrandIncorrect_Enter()
    HideBrandPlaceholder
End Sub

Private Sub txtBrandIncorrect_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtBrandIncorrect.Text)) = 0 Then
        ShowBrandPlaceholder
    End If
End Sub

' -- Placeholder helpers for txtPageRange --
Private Sub ShowPageRangePlaceholder()
    If txtPageRange Is Nothing Then Exit Sub
    If Len(Trim(txtPageRange.Text)) = 0 Or pageRangePlaceholderActive Then
        txtPageRange.Text = "e.g. 1,3,5-8"
        txtPageRange.ForeColor = &HC0C0C0  ' light grey
        pageRangePlaceholderActive = True
    End If
End Sub

Private Sub HidePageRangePlaceholder()
    If pageRangePlaceholderActive Then
        txtPageRange.Text = ""
        txtPageRange.ForeColor = &H0  ' black
        pageRangePlaceholderActive = False
    End If
End Sub

Private Function GetPageRangeText() As String
    If pageRangePlaceholderActive Then
        GetPageRangeText = ""
    Else
        GetPageRangeText = Trim(txtPageRange.Text)
    End If
End Function

Private Sub txtPageRange_Enter()
    HidePageRangePlaceholder
End Sub

Private Sub txtPageRange_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtPageRange.Text)) = 0 Then
        ShowPageRangePlaceholder
    End If
End Sub

Private Sub btnSaveBrands_Click()
    Dim brandFile As String
    brandFile = GetBrandRulesPath()

    ' Ensure directory exists (recursive, handles nested paths)
    Dim brandDir As String
    brandDir = GetParentDirectory(brandFile)
    If Len(brandDir) > 0 Then
        EnsureDirectoryExists brandDir
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
    ' Open a file picker instead of assuming the default path exists
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(3)  ' msoFileDialogFilePicker = 3
    If Err.Number <> 0 Then
        ' FileDialog not available -- fall back to default path
        Err.Clear
        On Error GoTo 0
        Dim fallbackPath As String
        fallbackPath = GetBrandRulesPath()
        If Dir(fallbackPath) = "" Then
            MsgBox "No saved brand rules found at:" & vbCrLf & fallbackPath, _
                   vbExclamation, "Brand Rules"
            Exit Sub
        End If
        LoadBrandRulesFromPath fallbackPath
        Exit Sub
    End If
    On Error GoTo 0

    With fd
        .Title = "Load Brand Rules"
        .AllowMultiSelect = False
        On Error Resume Next
        .Filters.Clear
        .Filters.Add "Text Files", "*.txt"
        .Filters.Add "All Files", "*.*"
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

        If .Show = -1 Then
            LoadBrandRulesFromPath CStr(.SelectedItems(1))
        End If
    End With
End Sub

Private Sub LoadBrandRulesFromPath(ByVal filePath As String)
    Dim loadResult As Boolean
    On Error Resume Next
    loadResult = Application.Run("Rules_Brands.LoadBrandRules", filePath)
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
        MsgBox "Brand rules loaded from:" & vbCrLf & filePath, vbInformation, "Brand Rules"
    Else
        MsgBox "Brand rules file could not be read:" & vbCrLf & filePath, _
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
