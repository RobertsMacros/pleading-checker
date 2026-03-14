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
' so that no .frx binary file is needed.
'
' Custom Rules: unified model replacing old Brand Rules and
' Custom Term Whitelist. Each rule has Enabled, Correct, and
' Incorrect Variants. Data is stored in Rules_Brands module.
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
Private WithEvents btnAddRule       As MSForms.CommandButton
Private WithEvents btnRemoveRule    As MSForms.CommandButton
Private WithEvents btnEditRule      As MSForms.CommandButton
Private WithEvents btnSaveRules     As MSForms.CommandButton
Private WithEvents btnLoadRules     As MSForms.CommandButton
Private WithEvents btnSortCorrect   As MSForms.CommandButton
Private WithEvents btnSortVariants  As MSForms.CommandButton
Private WithEvents btnSortDefault   As MSForms.CommandButton

Private fraRules        As MSForms.Frame
Private WithEvents txtPageRange As MSForms.TextBox
Private lstCustomRules  As MSForms.ListBox
Private txtRuleCorrect  As MSForms.TextBox
Private WithEvents txtRuleVariants As MSForms.TextBox
Private chkAddComments  As MSForms.CheckBox
Private chkTrackedChanges As MSForms.CheckBox
Private cboSpelling     As MSForms.ComboBox
Private cboQuoteNesting As MSForms.ComboBox
Private cboSmartQuotes  As MSForms.ComboBox
Private cboDateFormat   As MSForms.ComboBox
Private cboNonEngTerms  As MSForms.ComboBox
Private cboTermFormat   As MSForms.ComboBox
Private cboTermQuotes   As MSForms.ComboBox
Private cboSpaceStyle   As MSForms.ComboBox
Private lblStatus       As MSForms.Label

Private lastResults     As Collection
Private targetDoc       As Document
Private editingRuleIndex As Long     ' -1 = not editing; >= 0 = list index being edited
Private variantsPlaceholderActive As Boolean
Private pageRangePlaceholderActive As Boolean

' Custom rules data: parallel arrays for the single source of truth
Private crEnabled()     As Boolean   ' Whether rule is active at runtime
Private crCorrect()     As String    ' Correct form
Private crVariants()    As String    ' Comma-separated incorrect variants
Private crCount         As Long      ' Number of custom rules
Private crSortMode      As Long      ' 0=insertion, 1=correct, 2=variants
Private crSortOrder()   As Long      ' Indices into cr* arrays for display

' ============================================================
'  FORM INITIALISATION -- creates all controls at runtime
' ============================================================
Private Sub UserForm_Initialize()
    editingRuleIndex = -1
    variantsPlaceholderActive = False
    pageRangePlaceholderActive = False
    crCount = 0
    crSortMode = 0

    Dim lbl As MSForms.Label
    Dim yPos As Single

    ' -- Overall form padding ----------------------------------
    Const PAD As Single = 10
    Const FULL_W As Single = 680
    Const BTN_W As Single = 72
    Const BTN_H As Single = 20
    Const TXT_H As Single = 20
    Const CHK_H As Single = 16
    Const LBL_H As Single = 14
    Const SEC_GAP As Single = 6
    Const ITEM_GAP As Single = 2

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
    '  ROW 3: Left column (Page Range + Custom Rules)
    '         Right column (Options)
    ' ==========================================================
    Dim colLeft As Single
    Dim colRight As Single
    Dim leftW As Single
    Dim rightW As Single
    colLeft = PAD
    leftW = FULL_W * 0.56
    colRight = PAD + leftW + SEC_GAP
    rightW = FULL_W - leftW - SEC_GAP
    Dim row3Top As Single
    row3Top = yPos
    Dim cboW As Single
    cboW = rightW - 92

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

    ' ---- LEFT COLUMN: Custom Rules (unified section) ----
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblCustomRulesHeader")
    With lbl
        .Caption = "Custom Rules"
        .Left = colLeft: .Top = yPos: .Width = 80: .Height = LBL_H
        .Font.Size = 9: .Font.Bold = True
    End With

    ' Sort buttons inline with header
    Dim sortX As Single
    sortX = colLeft + 84
    Set btnSortDefault = Me.Controls.Add("Forms.CommandButton.1", "btnSortDefault")
    With btnSortDefault
        .Caption = "#"
        .Left = sortX: .Top = yPos - 1: .Width = 20: .Height = 16
        .Font.Size = 7
    End With
    Set btnSortCorrect = Me.Controls.Add("Forms.CommandButton.1", "btnSortCorrect")
    With btnSortCorrect
        .Caption = "A-Z"
        .Left = sortX + 22: .Top = yPos - 1: .Width = 28: .Height = 16
        .Font.Size = 7
    End With
    Set btnSortVariants = Me.Controls.Add("Forms.CommandButton.1", "btnSortVariants")
    With btnSortVariants
        .Caption = "Var"
        .Left = sortX + 52: .Top = yPos - 1: .Width = 28: .Height = 16
        .Font.Size = 7
    End With

    yPos = yPos + LBL_H + ITEM_GAP

    ' Custom rules listbox (multi-column table-like display)
    Dim ruleListW As Single
    ruleListW = leftW - BTN_W - ITEM_GAP - 4

    Set lstCustomRules = Me.Controls.Add("Forms.ListBox.1", "lstCustomRules")
    With lstCustomRules
        .Left = colLeft: .Top = yPos
        .Width = ruleListW: .Height = 72
        .Font.Size = 7.5
        .Font.Name = "Consolas"
        .ColumnCount = 4
        .ColumnWidths = "18;18;" & CStr(CLng(ruleListW * 0.35)) & ";" & CStr(CLng(ruleListW * 0.45))
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    ' Action buttons (right of list, stacked)
    Dim btnX As Single
    btnX = colLeft + ruleListW + ITEM_GAP
    Dim ruleBtnY As Single
    ruleBtnY = yPos

    Set btnAddRule = Me.Controls.Add("Forms.CommandButton.1", "btnAddRule")
    With btnAddRule
        .Caption = "Add"
        .Left = btnX: .Top = ruleBtnY: .Width = BTN_W: .Height = BTN_H
        .Font.Size = 7.5
    End With
    ruleBtnY = ruleBtnY + BTN_H + 1

    Set btnEditRule = Me.Controls.Add("Forms.CommandButton.1", "btnEditRule")
    With btnEditRule
        .Caption = "Edit"
        .Left = btnX: .Top = ruleBtnY: .Width = BTN_W / 2 - 1: .Height = BTN_H
        .Font.Size = 7
    End With

    Set btnRemoveRule = Me.Controls.Add("Forms.CommandButton.1", "btnRemoveRule")
    With btnRemoveRule
        .Caption = "Remove"
        .Left = btnX + BTN_W / 2 + 1: .Top = ruleBtnY: .Width = BTN_W / 2 - 1: .Height = BTN_H
        .Font.Size = 7
    End With
    ruleBtnY = ruleBtnY + BTN_H + 1

    Set btnSaveRules = Me.Controls.Add("Forms.CommandButton.1", "btnSaveRules")
    With btnSaveRules
        .Caption = "Save"
        .Left = btnX: .Top = ruleBtnY: .Width = BTN_W / 2 - 1: .Height = BTN_H
        .Font.Size = 7
    End With

    Set btnLoadRules = Me.Controls.Add("Forms.CommandButton.1", "btnLoadRules")
    With btnLoadRules
        .Caption = "Load"
        .Left = btnX + BTN_W / 2 + 1: .Top = ruleBtnY: .Width = BTN_W / 2 - 1: .Height = BTN_H
        .Font.Size = 7
    End With

    yPos = yPos + lstCustomRules.Height + ITEM_GAP

    ' Input row: Correct + Incorrect Variants
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblCorrectForm")
    With lbl
        .Caption = "Correct:"
        .Left = colLeft: .Top = yPos + 2: .Width = 42: .Height = LBL_H
        .Font.Size = 7.5
    End With

    Set txtRuleCorrect = Me.Controls.Add("Forms.TextBox.1", "txtRuleCorrect")
    With txtRuleCorrect
        .Left = colLeft + 42: .Top = yPos: .Width = 90: .Height = TXT_H
        .Font.Size = 7.5
    End With

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblIncorrectVars")
    With lbl
        .Caption = "Variants:"
        .Left = colLeft + 136: .Top = yPos + 2: .Width = 42: .Height = LBL_H
        .Font.Size = 7.5
    End With

    Set txtRuleVariants = Me.Controls.Add("Forms.TextBox.1", "txtRuleVariants")
    With txtRuleVariants
        .Left = colLeft + 178: .Top = yPos: .Width = ruleListW - 178 + colLeft: .Height = TXT_H
        .Font.Size = 7.5
    End With
    ShowVariantsPlaceholder

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
        .Left = colRight: .Top = optY: .Width = rightW: .Height = CHK_H
        .Value = True
        .Font.Size = 7.5
    End With
    optY = optY + CHK_H + ITEM_GAP

    Set chkTrackedChanges = Me.Controls.Add("Forms.CheckBox.1", "chkTrackedChanges")
    With chkTrackedChanges
        .Caption = "Tracked changes"
        .Left = colRight: .Top = optY: .Width = rightW: .Height = CHK_H
        .Value = True
        .Font.Size = 7.5
    End With
    optY = optY + CHK_H + ITEM_GAP

    ' Spelling mode
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblSpellingMode")
    With lbl
        .Caption = "Spelling:"
        .Left = colRight: .Top = optY + 2: .Width = 88: .Height = LBL_H
        .Font.Size = 7.5
    End With
    Set cboSpelling = Me.Controls.Add("Forms.ComboBox.1", "cboSpelling")
    With cboSpelling
        .Left = colRight + 90: .Top = optY: .Width = cboW: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "UK"
        .AddItem "US"
        .ListIndex = 0
        .Font.Size = 7.5
    End With
    optY = optY + TXT_H + ITEM_GAP

    ' Primary quotation marks
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblQuoteNesting")
    With lbl
        .Caption = "Primary quotation marks:"
        .Left = colRight: .Top = optY + 2: .Width = 88: .Height = LBL_H
        .Font.Size = 7
    End With
    Set cboQuoteNesting = Me.Controls.Add("Forms.ComboBox.1", "cboQuoteNesting")
    With cboQuoteNesting
        .Left = colRight + 90: .Top = optY: .Width = cboW: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "Single"
        .AddItem "Double"
        .ListIndex = 0
        .Font.Size = 7.5
    End With
    optY = optY + TXT_H + ITEM_GAP

    ' Smart quotes
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblSmartQuotes")
    With lbl
        .Caption = "Smart quotes:"
        .Left = colRight: .Top = optY + 2: .Width = 88: .Height = LBL_H
        .Font.Size = 7.5
    End With
    Set cboSmartQuotes = Me.Controls.Add("Forms.ComboBox.1", "cboSmartQuotes")
    With cboSmartQuotes
        .Left = colRight + 90: .Top = optY: .Width = cboW: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "Smart"
        .AddItem "Straight"
        .ListIndex = 0
        .Font.Size = 7.5
    End With
    optY = optY + TXT_H + ITEM_GAP

    ' Date format
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblDateFormat")
    With lbl
        .Caption = "Date format:"
        .Left = colRight: .Top = optY + 2: .Width = 88: .Height = LBL_H
        .Font.Size = 7.5
    End With
    Set cboDateFormat = Me.Controls.Add("Forms.ComboBox.1", "cboDateFormat")
    With cboDateFormat
        .Left = colRight + 90: .Top = optY: .Width = cboW + 46: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "UK (e.g. 14 March 2026 / 14/03/2026)"
        .AddItem "US (e.g. March 14, 2026 / 03/14/2026)"
        .ListIndex = 0
        .Font.Size = 7.5
    End With
    optY = optY + TXT_H + ITEM_GAP

    ' Non-English Terms
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblNonEngTerms")
    With lbl
        .Caption = "Non-English Terms:"
        .Left = colRight: .Top = optY + 2: .Width = 88: .Height = LBL_H
        .Font.Size = 7.5
    End With
    Set cboNonEngTerms = Me.Controls.Add("Forms.ComboBox.1", "cboNonEngTerms")
    With cboNonEngTerms
        .Left = colRight + 90: .Top = optY: .Width = cboW: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "Italics"
        .AddItem "Regular text"
        .ListIndex = 0
        .Font.Size = 7.5
    End With
    optY = optY + TXT_H + ITEM_GAP

    ' After full stop
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblSpaceStyle")
    With lbl
        .Caption = "After full stop:"
        .Left = colRight: .Top = optY + 2: .Width = 88: .Height = LBL_H
        .Font.Size = 7.5
    End With
    Set cboSpaceStyle = Me.Controls.Add("Forms.ComboBox.1", "cboSpaceStyle")
    With cboSpaceStyle
        .Left = colRight + 90: .Top = optY: .Width = cboW: .Height = TXT_H
        .Style = fmStyleDropDownList
        .AddItem "One space"
        .AddItem "Two spaces"
        .ListIndex = 0
        .Font.Size = 7.5
    End With
    optY = optY + TXT_H + ITEM_GAP

    ' Defined terms (format + quotes)
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblDefinedTerms")
    With lbl
        .Caption = "Def. terms:"
        .Left = colRight: .Top = optY + 2: .Width = 56: .Height = LBL_H
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
        .Left = colRight + 128: .Top = optY + 2: .Width = 10: .Height = LBL_H
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

    ' -- Load custom rules from engine -------------------------
    LoadCustomRulesFromEngine

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

    Const COLS As Long = 4
    Const ROW_H As Single = 18
    Const COL_PAD As Single = 6

    Const MIN_USABLE_W As Single = 120
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
'  CUSTOM RULES DATA MODEL
'  Single source of truth: crEnabled(), crCorrect(), crVariants()
' ============================================================
Private Sub InitCustomRulesArrays(ByVal capacity As Long)
    If capacity < 1 Then capacity = 1
    ReDim crEnabled(0 To capacity - 1)
    ReDim crCorrect(0 To capacity - 1)
    ReDim crVariants(0 To capacity - 1)
    ReDim crSortOrder(0 To capacity - 1)
    crCount = 0
End Sub

Private Sub AddCustomRule(ByVal correct As String, ByVal variants As String, ByVal enabled As Boolean)
    If crCount = 0 Then
        InitCustomRulesArrays 16
    End If
    ' Grow arrays if needed
    If crCount > UBound(crCorrect) Then
        ReDim Preserve crEnabled(0 To crCount * 2)
        ReDim Preserve crCorrect(0 To crCount * 2)
        ReDim Preserve crVariants(0 To crCount * 2)
        ReDim Preserve crSortOrder(0 To crCount * 2)
    End If
    crCorrect(crCount) = correct
    crVariants(crCount) = variants
    crEnabled(crCount) = enabled
    crCount = crCount + 1
    RebuildSortOrder
End Sub

Private Sub RemoveCustomRule(ByVal idx As Long)
    If idx < 0 Or idx >= crCount Then Exit Sub
    Dim j As Long
    For j = idx To crCount - 2
        crCorrect(j) = crCorrect(j + 1)
        crVariants(j) = crVariants(j + 1)
        crEnabled(j) = crEnabled(j + 1)
    Next j
    crCount = crCount - 1
    RebuildSortOrder
End Sub

Private Sub UpdateCustomRule(ByVal idx As Long, ByVal correct As String, ByVal variants As String)
    If idx < 0 Or idx >= crCount Then Exit Sub
    crCorrect(idx) = correct
    crVariants(idx) = variants
    RebuildSortOrder
End Sub

Private Sub ToggleCustomRuleEnabled(ByVal idx As Long)
    If idx < 0 Or idx >= crCount Then Exit Sub
    crEnabled(idx) = Not crEnabled(idx)
End Sub

Private Sub RebuildSortOrder()
    If crCount = 0 Then
        ReDim crSortOrder(0 To 0)
        Exit Sub
    End If
    ReDim crSortOrder(0 To crCount - 1)
    Dim i As Long
    For i = 0 To crCount - 1
        crSortOrder(i) = i
    Next i
    If crSortMode = 1 Then
        SortOrderByField 1  ' sort by correct
    ElseIf crSortMode = 2 Then
        SortOrderByField 2  ' sort by variants
    End If
End Sub

Private Sub SortOrderByField(ByVal field As Long)
    ' Simple insertion sort on crSortOrder by the chosen field
    Dim i As Long, j As Long, tmp As Long
    Dim valI As String, valJ As String
    For i = 1 To crCount - 1
        tmp = crSortOrder(i)
        If field = 1 Then
            valI = LCase$(crCorrect(tmp))
        Else
            valI = LCase$(crVariants(tmp))
        End If
        j = i - 1
        Do While j >= 0
            If field = 1 Then
                valJ = LCase$(crCorrect(crSortOrder(j)))
            Else
                valJ = LCase$(crVariants(crSortOrder(j)))
            End If
            If valJ <= valI Then Exit Do
            crSortOrder(j + 1) = crSortOrder(j)
            j = j - 1
        Loop
        crSortOrder(j + 1) = tmp
    Next i
End Sub

' Map a display-list index to a data-array index
Private Function DisplayToDataIndex(ByVal displayIdx As Long) As Long
    If displayIdx < 0 Or displayIdx >= crCount Then
        DisplayToDataIndex = -1
    Else
        DisplayToDataIndex = crSortOrder(displayIdx)
    End If
End Function

' ============================================================
'  LOAD CUSTOM RULES FROM ENGINE (Rules_Brands module)
' ============================================================
Private Sub LoadCustomRulesFromEngine()
    InitCustomRulesArrays 16
    On Error Resume Next
    Dim brands As Object
    Set brands = Application.Run("Rules_Brands.GetBrandRules")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        RefreshCustomRulesList
        Exit Sub
    End If
    On Error GoTo 0
    If brands Is Nothing Then
        RefreshCustomRulesList
        Exit Sub
    End If
    Dim bKey As Variant
    For Each bKey In brands.keys
        AddCustomRule CStr(bKey), CStr(brands(bKey)), True
    Next bKey
    RefreshCustomRulesList
End Sub

' ============================================================
'  SYNC CUSTOM RULES BACK TO ENGINE (before run)
' ============================================================
Private Sub SyncCustomRulesToEngine()
    ' Clear existing brand rules and rebuild from our data model
    On Error Resume Next
    Dim oldBrands As Object
    Set oldBrands = Application.Run("Rules_Brands.GetBrandRules")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Remove all existing rules
    If Not oldBrands Is Nothing Then
        Dim oldKeys As Variant
        oldKeys = oldBrands.keys
        Dim m As Long
        For m = UBound(oldKeys) To LBound(oldKeys) Step -1
            On Error Resume Next
            Application.Run "Rules_Brands.RemoveBrandRule", CStr(oldKeys(m))
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        Next m
    End If

    ' Add back only enabled rules
    Dim n As Long
    For n = 0 To crCount - 1
        If crEnabled(n) Then
            On Error Resume Next
            Application.Run "Rules_Brands.AddBrandRule", crCorrect(n), crVariants(n)
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        End If
    Next n
End Sub

' ============================================================
'  REFRESH CUSTOM RULES LISTBOX
' ============================================================
Private Sub RefreshCustomRulesList()
    lstCustomRules.Clear
    If crCount = 0 Then Exit Sub
    Dim i As Long
    Dim di As Long
    Dim enabledMark As String
    For i = 0 To crCount - 1
        di = crSortOrder(i)
        If crEnabled(di) Then enabledMark = ChrW$(9745) Else enabledMark = ChrW$(9744)
        lstCustomRules.AddItem ""
        lstCustomRules.List(i, 0) = enabledMark
        lstCustomRules.List(i, 1) = CStr(di + 1)
        lstCustomRules.List(i, 2) = crCorrect(di)
        lstCustomRules.List(i, 3) = crVariants(di)
    Next i
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

    ' Sync custom rules to engine (only enabled rules)
    SyncCustomRulesToEngine

    ' Set page range from flexible input (ignore placeholder text)
    PleadingsEngine.SetPageRangeFromString GetPageRangeText()

    ' Set mode toggles from dropdowns
    If cboSpelling.ListIndex = 1 Then
        PleadingsEngine.SetSpellingMode "US"
    Else
        PleadingsEngine.SetSpellingMode "UK"
    End If

    If cboQuoteNesting.ListIndex = 1 Then
        PleadingsEngine.SetQuoteNesting "DOUBLE"
    Else
        PleadingsEngine.SetQuoteNesting "SINGLE"
    End If

    If cboSmartQuotes.ListIndex = 1 Then
        PleadingsEngine.SetSmartQuotePref "STRAIGHT"
    Else
        PleadingsEngine.SetSmartQuotePref "SMART"
    End If

    If cboDateFormat.ListIndex = 1 Then
        PleadingsEngine.SetDateFormatPref "US"
    Else
        PleadingsEngine.SetDateFormatPref "UK"
    End If

    If cboNonEngTerms.ListIndex = 1 Then
        PleadingsEngine.SetNonEngTermPref "REGULAR"
    Else
        PleadingsEngine.SetNonEngTermPref "ITALICS"
    End If

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

    If Len(reportPath) = 0 Then
        reportPath = GetTempReportPath(sep)
    End If

    Dim reportDir As String
    reportDir = GetParentDirectory(reportPath)
    If Len(reportDir) > 0 Then
        EnsureDirectoryExists reportDir
    End If

    lblStatus.Caption = "Exporting report..."
    Me.Repaint
    DoEvents

    Dim reportSummary As String
    reportSummary = PleadingsEngine.GenerateReport(lastResults, reportPath, targetDoc)

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

    msg = msg & vbCrLf & vbCrLf & reportSummary

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
'  CUSTOM RULES: ADD
' ============================================================
Private Sub btnAddRule_Click()
    Dim correctForm As String
    Dim incorrectVars As String
    correctForm = Trim$(txtRuleCorrect.Text)
    incorrectVars = GetVariantsText()

    If Len(correctForm) = 0 Then
        MsgBox "Enter the correct form.", vbExclamation, "Custom Rules"
        txtRuleCorrect.SetFocus
        Exit Sub
    End If

    If Len(incorrectVars) = 0 Then
        MsgBox "Enter at least one incorrect variant.", vbExclamation, "Custom Rules"
        txtRuleVariants.SetFocus
        Exit Sub
    End If

    ' Normalise variants
    incorrectVars = NormaliseVariants(incorrectVars)
    If Len(incorrectVars) = 0 Then
        MsgBox "Enter at least one incorrect variant.", vbExclamation, "Custom Rules"
        txtRuleVariants.SetFocus
        Exit Sub
    End If

    If editingRuleIndex >= 0 Then
        ' Update existing rule
        Dim dataIdx As Long
        dataIdx = DisplayToDataIndex(editingRuleIndex)
        If dataIdx >= 0 Then
            UpdateCustomRule dataIdx, correctForm, incorrectVars
        End If
        editingRuleIndex = -1
        btnAddRule.Caption = "Add"
    Else
        ' Add new rule (enabled by default)
        AddCustomRule correctForm, incorrectVars, True
    End If

    txtRuleCorrect.Text = ""
    ClearVariants
    ShowVariantsPlaceholder
    RefreshCustomRulesList
End Sub

' ============================================================
'  CUSTOM RULES: REMOVE
' ============================================================
Private Sub btnRemoveRule_Click()
    If lstCustomRules.ListIndex < 0 Then
        MsgBox "Select a rule to remove.", vbExclamation, "Custom Rules"
        Exit Sub
    End If

    Dim dataIdx As Long
    dataIdx = DisplayToDataIndex(lstCustomRules.ListIndex)
    If dataIdx >= 0 Then
        RemoveCustomRule dataIdx
    End If

    ' Cancel any edit in progress
    If editingRuleIndex >= 0 Then
        editingRuleIndex = -1
        btnAddRule.Caption = "Add"
        txtRuleCorrect.Text = ""
        ClearVariants
        ShowVariantsPlaceholder
    End If

    RefreshCustomRulesList
End Sub

' ============================================================
'  CUSTOM RULES: EDIT
' ============================================================
Private Sub btnEditRule_Click()
    If lstCustomRules.ListIndex < 0 Then
        MsgBox "Select a rule to edit.", vbInformation, "Custom Rules"
        Exit Sub
    End If

    Dim dataIdx As Long
    dataIdx = DisplayToDataIndex(lstCustomRules.ListIndex)
    If dataIdx < 0 Then Exit Sub

    editingRuleIndex = lstCustomRules.ListIndex
    txtRuleCorrect.Text = crCorrect(dataIdx)
    HideVariantsPlaceholder
    txtRuleVariants.Text = crVariants(dataIdx)
    btnAddRule.Caption = "Save Edit"
End Sub

' ============================================================
'  CUSTOM RULES: TOGGLE ENABLED (double-click on list)
' ============================================================
' Note: MSForms.ListBox does not have a DblClick WithEvents in
' the same way. We use a workaround: user selects row and we
' provide a toggle via the checkbox column click.
' Since we cannot intercept individual column clicks in a
' standard MSForms.ListBox, we toggle enabled state when the
' user double-clicks (handled by the list control's built-in
' DblClick event if available, or we add a Toggle button).
' For simplicity, we repurpose the # column header button.

' ============================================================
'  CUSTOM RULES: SORT BUTTONS
' ============================================================
Private Sub btnSortDefault_Click()
    crSortMode = 0
    RebuildSortOrder
    RefreshCustomRulesList
End Sub

Private Sub btnSortCorrect_Click()
    crSortMode = 1
    RebuildSortOrder
    RefreshCustomRulesList
End Sub

Private Sub btnSortVariants_Click()
    crSortMode = 2
    RebuildSortOrder
    RefreshCustomRulesList
End Sub

' ============================================================
'  CUSTOM RULES: SAVE
' ============================================================
Private Sub btnSaveRules_Click()
    Dim rulesFile As String
    rulesFile = GetCustomRulesPath()

    Dim rulesDir As String
    rulesDir = GetParentDirectory(rulesFile)
    If Len(rulesDir) > 0 Then
        EnsureDirectoryExists rulesDir
    End If

    ' Save all rules (enabled and disabled) with enabled flag
    Dim fileNum As Integer
    fileNum = FreeFile
    On Error GoTo SaveFail
    Open rulesFile For Output As #fileNum
    Dim s As Long
    For s = 0 To crCount - 1
        Dim prefix As String
        If crEnabled(s) Then prefix = "+" Else prefix = "-"
        Print #fileNum, prefix & crCorrect(s) & "=" & crVariants(s)
    Next s
    Close #fileNum

    MsgBox "Custom rules saved to:" & vbCrLf & rulesFile, vbInformation, "Custom Rules"
    Exit Sub

SaveFail:
    On Error Resume Next
    Close #fileNum
    MsgBox "Failed to save custom rules to:" & vbCrLf & rulesFile & vbCrLf & _
           "Error: " & Err.Description, vbExclamation, "Custom Rules"
    Err.Clear
    On Error GoTo 0
End Sub

' ============================================================
'  CUSTOM RULES: LOAD
' ============================================================
Private Sub btnLoadRules_Click()
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(3)  ' msoFileDialogFilePicker
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Dim fallbackPath As String
        fallbackPath = GetCustomRulesPath()
        If Dir(fallbackPath) = "" Then
            MsgBox "No saved custom rules found at:" & vbCrLf & fallbackPath, _
                   vbExclamation, "Custom Rules"
            Exit Sub
        End If
        LoadCustomRulesFromFile fallbackPath
        Exit Sub
    End If
    On Error GoTo 0

    With fd
        .Title = "Load Custom Rules"
        .AllowMultiSelect = False
        On Error Resume Next
        .Filters.Clear
        .Filters.Add "Text Files", "*.txt"
        .Filters.Add "All Files", "*.*"
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0

        If .Show = -1 Then
            LoadCustomRulesFromFile CStr(.SelectedItems(1))
        End If
    End With
End Sub

Private Sub LoadCustomRulesFromFile(ByVal filePath As String)
    Dim fileNum As Integer
    Dim lineText As String
    Dim eqPos As Long
    Dim correct As String
    Dim variants As String
    Dim enabled As Boolean

    InitCustomRulesArrays 16

    fileNum = FreeFile
    On Error GoTo LoadFail
    Open filePath For Input As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim$(lineText)
        If Len(lineText) = 0 Then GoTo NextLoadLine
        If Left$(lineText, 1) = "#" Then GoTo NextLoadLine

        ' Parse enabled flag: + or - prefix
        enabled = True
        If Left$(lineText, 1) = "+" Then
            lineText = Mid$(lineText, 2)
        ElseIf Left$(lineText, 1) = "-" Then
            enabled = False
            lineText = Mid$(lineText, 2)
        End If

        eqPos = InStr(lineText, "=")
        If eqPos > 1 Then
            correct = Trim$(Left$(lineText, eqPos - 1))
            variants = Trim$(Mid$(lineText, eqPos + 1))
            If Len(correct) > 0 And Len(variants) > 0 Then
                AddCustomRule correct, variants, enabled
            End If
        End If

NextLoadLine:
    Loop

    Close #fileNum
    RefreshCustomRulesList

    ' Also sync enabled rules to engine
    SyncCustomRulesToEngine

    MsgBox "Custom rules loaded from:" & vbCrLf & filePath, vbInformation, "Custom Rules"
    Exit Sub

LoadFail:
    On Error Resume Next
    Close #fileNum
    MsgBox "Could not read custom rules from:" & vbCrLf & filePath & vbCrLf & _
           "Error: " & Err.Description, vbExclamation, "Custom Rules"
    Err.Clear
    On Error GoTo 0
    ' Fall back to engine defaults if nothing loaded
    If crCount = 0 Then LoadCustomRulesFromEngine
End Sub

' -- Helper: cross-platform custom rules file path --
Private Function GetCustomRulesPath() As String
    On Error Resume Next
    GetCustomRulesPath = Application.Run("Rules_Brands.GetDefaultBrandRulesPath")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Dim sep As String
        sep = Application.PathSeparator
        #If Mac Then
            GetCustomRulesPath = Environ("HOME") & sep & "Library" & sep & _
                                "Application Support" & sep & "PleadingsChecker" & sep & "brand_rules.txt"
        #Else
            GetCustomRulesPath = Environ("APPDATA") & sep & "PleadingsChecker" & sep & "brand_rules.txt"
        #End If
        Exit Function
    End If
    On Error GoTo 0
End Function

' ============================================================
'  VARIANTS TEXTBOX PLACEHOLDER HELPERS
' ============================================================
Private Sub ShowVariantsPlaceholder()
    If txtRuleVariants Is Nothing Then Exit Sub
    If Len(Trim$(txtRuleVariants.Text)) = 0 Or variantsPlaceholderActive Then
        txtRuleVariants.Text = "e.g. colour, color, colours"
        txtRuleVariants.ForeColor = &HC0C0C0  ' light grey
        variantsPlaceholderActive = True
    End If
End Sub

Private Sub HideVariantsPlaceholder()
    If variantsPlaceholderActive Then
        txtRuleVariants.Text = ""
        txtRuleVariants.ForeColor = &H0  ' black
        variantsPlaceholderActive = False
    End If
End Sub

Private Function GetVariantsText() As String
    If variantsPlaceholderActive Then
        GetVariantsText = ""
    Else
        GetVariantsText = Trim$(txtRuleVariants.Text)
    End If
End Function

Private Sub ClearVariants()
    txtRuleVariants.Text = ""
    txtRuleVariants.ForeColor = &H0
    variantsPlaceholderActive = False
End Sub

Private Sub txtRuleVariants_Enter()
    HideVariantsPlaceholder
End Sub

Private Sub txtRuleVariants_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim$(txtRuleVariants.Text)) = 0 Then
        ShowVariantsPlaceholder
    End If
End Sub

' Normalise comma-separated variants: trim each, remove blanks
Private Function NormaliseVariants(ByVal raw As String) As String
    Dim parts() As String
    parts = Split(raw, ",")
    Dim result As String
    Dim p As Long
    For p = LBound(parts) To UBound(parts)
        Dim item As String
        item = Trim$(parts(p))
        If Len(item) > 0 Then
            If Len(result) > 0 Then result = result & ", "
            result = result & item
        End If
    Next p
    NormaliseVariants = result
End Function

' ============================================================
'  PAGE RANGE PLACEHOLDER HELPERS
' ============================================================
Private Sub ShowPageRangePlaceholder()
    If txtPageRange Is Nothing Then Exit Sub
    If Len(Trim$(txtPageRange.Text)) = 0 Or pageRangePlaceholderActive Then
        txtPageRange.Text = "e.g. 1,3,5-8,9:30"
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
        GetPageRangeText = Trim$(txtPageRange.Text)
    End If
End Function

Private Sub txtPageRange_Enter()
    HidePageRangePlaceholder
End Sub

Private Sub txtPageRange_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim$(txtPageRange.Text)) = 0 Then
        ShowPageRangePlaceholder
    End If
End Sub

' ============================================================
'  CLOSE BUTTON
' ============================================================
Private Sub btnClose_Click()
    Unload Me
End Sub
