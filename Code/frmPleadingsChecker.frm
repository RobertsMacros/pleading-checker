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
' Custom Rules: unified model. Each rule has Correct and
' Incorrect Variants. Persistence via custom_rules.txt.
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

' Header click handlers for column sorting (clsHeaderClick class)
Private mHdrNum      As clsHeaderClick
Private mHdrCorrect  As clsHeaderClick
Private mHdrVariants As clsHeaderClick

Private fraRules        As MSForms.Frame
Private WithEvents txtPageRange As MSForms.TextBox
Private lstCustomRules  As MSForms.ListBox
Private WithEvents txtRuleCorrect As MSForms.TextBox
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

' Placeholder constants and flags
Private Const PAGE_RANGE_PLACEHOLDER As String = "e.g. 1,3,5-8,9:30"
Private mPageRangeShowingPlaceholder As Boolean
Private Const VARIANTS_PLACEHOLDER As String = "e.g. colour, color, colours"
Private mVariantsShowingPlaceholder As Boolean
Private Const CORRECT_PLACEHOLDER As String = "e.g. colour"
Private mCorrectShowingPlaceholder As Boolean

' Custom rules data: parallel arrays for the single source of truth
Private crCorrect()     As String    ' Correct form
Private crVariants()    As String    ' Comma-separated incorrect variants
Private crInsertSeq()   As Long      ' Insertion-order sequence number
Private crCount         As Long      ' Number of custom rules
Private crNextSeq       As Long      ' Next sequence number to assign
Private crSortMode      As Long      ' 0=insertion, 1=correct, 2=variants
Private crSortAscending As Boolean   ' True=ascending, False=descending
Private crSortOrder()   As Long      ' Indices into cr* arrays for display

' ============================================================
'  CONTROL-CREATION HELPERS
'  Reduce duplication in UserForm_Initialize.
' ============================================================
Private Function AddLabel(parent As Object, ByVal ctlName As String, _
                          ByVal cap As String, ByVal l As Single, _
                          ByVal t As Single, ByVal w As Single, _
                          ByVal h As Single, _
                          Optional ByVal fontSize As Single = 7.5, _
                          Optional ByVal isBold As Boolean = False) As MSForms.Label
    Set AddLabel = parent.Controls.Add("Forms.Label.1", ctlName)
    With AddLabel
        .Caption = cap
        .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Size = fontSize
        .Font.Bold = isBold
    End With
End Function

Private Function AddButton(parent As Object, ByVal ctlName As String, _
                           ByVal cap As String, ByVal l As Single, _
                           ByVal t As Single, ByVal w As Single, _
                           ByVal h As Single, _
                           Optional ByVal fontSize As Single = 7.5, _
                           Optional ByVal isBold As Boolean = False) As MSForms.CommandButton
    Set AddButton = parent.Controls.Add("Forms.CommandButton.1", ctlName)
    With AddButton
        .Caption = cap
        .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Size = fontSize
        .Font.Bold = isBold
    End With
End Function

Private Function AddTextBox(parent As Object, ByVal ctlName As String, _
                            ByVal l As Single, ByVal t As Single, _
                            ByVal w As Single, ByVal h As Single, _
                            Optional ByVal fontSize As Single = 7.5) As MSForms.TextBox
    Set AddTextBox = parent.Controls.Add("Forms.TextBox.1", ctlName)
    With AddTextBox
        .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Size = fontSize
        .Text = ""
    End With
End Function

Private Function AddCombo(parent As Object, ByVal ctlName As String, _
                          ByVal l As Single, ByVal t As Single, _
                          ByVal w As Single, ByVal h As Single, _
                          Optional ByVal fontSize As Single = 7.5) As MSForms.ComboBox
    Set AddCombo = parent.Controls.Add("Forms.ComboBox.1", ctlName)
    With AddCombo
        .Left = l: .Top = t: .Width = w: .Height = h
        .Style = fmStyleDropDownList
        .Font.Size = fontSize
    End With
End Function

Private Function AddCheckBox(parent As Object, ByVal ctlName As String, _
                             ByVal cap As String, ByVal l As Single, _
                             ByVal t As Single, ByVal w As Single, _
                             ByVal h As Single, _
                             Optional ByVal fontSize As Single = 7.5, _
                             Optional ByVal startVal As Boolean = True) As MSForms.CheckBox
    Set AddCheckBox = parent.Controls.Add("Forms.CheckBox.1", ctlName)
    With AddCheckBox
        .Caption = cap
        .Left = l: .Top = t: .Width = w: .Height = h
        .Value = startVal
        .Font.Size = fontSize
    End With
End Function

' Helper: create a "Label + ComboBox" option row and return the combo.
' Advances yPos by rowHeight + gap.
Private Function AddOptionRow(ByVal labelText As String, _
                              ByVal ctlName As String, _
                              ByRef yPos As Single, _
                              ByVal colRight As Single, _
                              ByVal lblOptW As Single, _
                              ByVal cboW As Single, _
                              ByVal rowH As Single, _
                              ByVal lblH As Single, _
                              ByVal gap As Single, _
                              items As Variant, _
                              Optional ByVal defaultIdx As Long = 0) As MSForms.ComboBox
    AddLabel Me, "lbl" & ctlName, labelText, colRight, yPos + 2, lblOptW, lblH
    Set AddOptionRow = AddCombo(Me, ctlName, colRight + lblOptW + 2, yPos, cboW, rowH)
    Dim v As Variant
    For Each v In items
        AddOptionRow.AddItem CStr(v)
    Next v
    AddOptionRow.ListIndex = defaultIdx
    yPos = yPos + rowH + gap
End Function

' ============================================================
'  FORM INITIALISATION -- creates all controls at runtime
' ============================================================
Private Sub UserForm_Initialize()
    editingRuleIndex = -1
    mVariantsShowingPlaceholder = False
    mCorrectShowingPlaceholder = False
    mPageRangeShowingPlaceholder = False
    crCount = 0
    crSortMode = 0
    crSortAscending = True

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
    AddLabel Me, "lblRulesHeader", "Rules", PAD, yPos, 40, LBL_H, 9, True
    Set btnSelectAll = AddButton(Me, "btnSelectAll", "Select All", PAD + 44, yPos - 1, 62, 18, 7)
    Set btnDeselectAll = AddButton(Me, "btnDeselectAll", "Deselect All", PAD + 44 + 64, yPos - 1, 62, 18, 7)

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
    Dim lblOptW As Single
    lblOptW = 82
    cboW = rightW - lblOptW - 2

    ' ---- LEFT COLUMN: Page Range ----
    AddLabel Me, "lblPageHeader", "Page Range", colLeft, yPos, 120, LBL_H, 9, True
    yPos = yPos + LBL_H + ITEM_GAP

    AddLabel Me, "lblPageRange", "Pages:", colLeft, yPos + 2, 36, LBL_H
    Set txtPageRange = AddTextBox(Me, "txtPageRange", colLeft + 36, yPos, leftW - 40, TXT_H)
    InitPlaceholder txtPageRange, PAGE_RANGE_PLACEHOLDER, mPageRangeShowingPlaceholder

    yPos = yPos + TXT_H + SEC_GAP

    ' ---- LEFT COLUMN: Custom Rules (unified section) ----
    AddLabel Me, "lblCustomRulesHeader", "Custom Rules", colLeft, yPos, 80, LBL_H, 9, True

    yPos = yPos + LBL_H + ITEM_GAP

    ' Custom rules table: column header row + listbox (3 columns)
    Dim ruleListW As Single
    ruleListW = leftW - BTN_W - ITEM_GAP - 4
    Dim colWNum As Single:     colWNum = 24
    Dim colWCorrect As Single: colWCorrect = CLng((ruleListW - colWNum) * 0.4)
    Dim colWVariants As Single: colWVariants = ruleListW - colWNum - colWCorrect
    Dim hdrH As Single:        hdrH = 14

    ' Clickable column header labels (wired via clsHeaderClick)
    Dim lblH As MSForms.Label
    Set lblH = AddLabel(Me, "lblHdrNum", " #", colLeft, yPos, colWNum, hdrH, 7, True)
    lblH.TextAlign = fmTextAlignCenter
    lblH.BackColor = RGB(230, 230, 230): lblH.BackStyle = fmBackStyleOpaque
    Set mHdrNum = New clsHeaderClick
    mHdrNum.Init Me, lblH, 0

    Set lblH = AddLabel(Me, "lblHdrCorrect", " Correct", colLeft + colWNum, yPos, colWCorrect, hdrH, 7, True)
    lblH.BackColor = RGB(198, 239, 206): lblH.BackStyle = fmBackStyleOpaque
    Set mHdrCorrect = New clsHeaderClick
    mHdrCorrect.Init Me, lblH, 1

    Set lblH = AddLabel(Me, "lblHdrVariants", " Incorrect Variants", colLeft + colWNum + colWCorrect, yPos, colWVariants, hdrH, 7, True)
    lblH.BackColor = RGB(255, 199, 206): lblH.BackStyle = fmBackStyleOpaque
    Set mHdrVariants = New clsHeaderClick
    mHdrVariants.Init Me, lblH, 2

    yPos = yPos + hdrH

    ' Custom rules listbox (3-column table-like display)
    Set lstCustomRules = Me.Controls.Add("Forms.ListBox.1", "lstCustomRules")
    With lstCustomRules
        .Left = colLeft: .Top = yPos
        .Width = ruleListW: .Height = 62
        .Font.Size = 7.5
        .Font.Name = "Consolas"
        .ColumnCount = 3
        .ColumnWidths = CStr(CLng(colWNum)) & ";" & _
                        CStr(CLng(colWCorrect)) & ";" & CStr(CLng(colWVariants))
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
    End With

    ' Action buttons (right of list, stacked)
    Dim btnX As Single
    btnX = colLeft + ruleListW + ITEM_GAP
    Dim ruleBtnY As Single
    ruleBtnY = yPos

    Set btnAddRule = AddButton(Me, "btnAddRule", "Add", btnX, ruleBtnY, BTN_W, BTN_H)
    ruleBtnY = ruleBtnY + BTN_H + 1
    Set btnEditRule = AddButton(Me, "btnEditRule", "Edit", btnX, ruleBtnY, BTN_W / 2 - 1, BTN_H, 7)
    Set btnRemoveRule = AddButton(Me, "btnRemoveRule", "Remove", btnX + BTN_W / 2 + 1, ruleBtnY, BTN_W / 2 - 1, BTN_H, 7)
    ruleBtnY = ruleBtnY + BTN_H + 1
    Set btnSaveRules = AddButton(Me, "btnSaveRules", "Save", btnX, ruleBtnY, BTN_W / 2 - 1, BTN_H, 7)
    Set btnLoadRules = AddButton(Me, "btnLoadRules", "Load", btnX + BTN_W / 2 + 1, ruleBtnY, BTN_W / 2 - 1, BTN_H, 7)

    yPos = yPos + lstCustomRules.Height + ITEM_GAP

    ' Input row: Correct + Incorrect Variants
    AddLabel Me, "lblCorrectForm", "Correct:", colLeft, yPos + 2, 42, LBL_H
    Set txtRuleCorrect = AddTextBox(Me, "txtRuleCorrect", colLeft + 42, yPos, 90, TXT_H)
    InitPlaceholder txtRuleCorrect, CORRECT_PLACEHOLDER, mCorrectShowingPlaceholder

    AddLabel Me, "lblIncorrectVars", "Variants:", colLeft + 136, yPos + 2, 42, LBL_H
    Set txtRuleVariants = AddTextBox(Me, "txtRuleVariants", colLeft + 178, yPos, ruleListW - 178 + colLeft, TXT_H)
    InitPlaceholder txtRuleVariants, VARIANTS_PLACEHOLDER, mVariantsShowingPlaceholder

    Dim leftBottomY As Single
    leftBottomY = yPos + TXT_H

    ' ---- RIGHT COLUMN: Options (starting from row3Top) ----
    Dim optY As Single
    optY = row3Top

    AddLabel Me, "lblOptionsHeader", "Options", colRight, optY, 120, LBL_H, 9, True
    optY = optY + LBL_H + ITEM_GAP

    Set chkAddComments = AddCheckBox(Me, "chkAddComments", "Add comments", colRight, optY, rightW, CHK_H)
    optY = optY + CHK_H + ITEM_GAP

    Set chkTrackedChanges = AddCheckBox(Me, "chkTrackedChanges", "Tracked changes", colRight, optY, rightW, CHK_H)
    optY = optY + CHK_H + ITEM_GAP

    Set cboSpelling = AddOptionRow("Spelling:", "cboSpelling", optY, colRight, lblOptW, cboW, TXT_H, LBL_H, ITEM_GAP, Array("UK", "US"))
    Set cboQuoteNesting = AddOptionRow("Primary quotes:", "cboQuoteNesting", optY, colRight, lblOptW, cboW, TXT_H, LBL_H, ITEM_GAP, Array("Single", "Double"))
    Set cboSmartQuotes = AddOptionRow("Smart quotes:", "cboSmartQuotes", optY, colRight, lblOptW, cboW, TXT_H, LBL_H, ITEM_GAP, Array("Smart", "Straight"))
    Set cboDateFormat = AddOptionRow("Date format:", "cboDateFormat", optY, colRight, lblOptW, cboW, TXT_H, LBL_H, ITEM_GAP, Array("UK (14 March 2026 / 14/03/2026)", "US (March 14, 2026 / 03/14/2026)"))
    Set cboNonEngTerms = AddOptionRow("Non-English terms:", "cboNonEngTerms", optY, colRight, lblOptW, cboW, TXT_H, LBL_H, ITEM_GAP, Array("Italics", "Regular text"))
    Set cboSpaceStyle = AddOptionRow("After full stop:", "cboSpaceStyle", optY, colRight, lblOptW, cboW, TXT_H, LBL_H, ITEM_GAP, Array("One space", "Two spaces"))

    ' Defined terms formatting pair
    Dim dtLblW As Single: dtLblW = 52
    Dim dtCboW As Single: dtCboW = (rightW - dtLblW - 14) / 2
    AddLabel Me, "lblDefinedTerms", "Def. terms:", colRight, optY + 2, dtLblW, LBL_H
    Set cboTermFormat = AddCombo(Me, "cboTermFormat", colRight + dtLblW + 2, optY, dtCboW, TXT_H)
    Dim tfItem As Variant
    For Each tfItem In Array("Bold", "Bold Italics", "Italics", "None")
        cboTermFormat.AddItem CStr(tfItem)
    Next tfItem
    cboTermFormat.ListIndex = 0

    AddLabel Me, "lblTermAnd", "+", colRight + dtLblW + dtCboW + 4, optY + 2, 10, LBL_H
    Set cboTermQuotes = AddCombo(Me, "cboTermQuotes", colRight + dtLblW + dtCboW + 14, optY, dtCboW, TXT_H)
    cboTermQuotes.AddItem "Single quotes"
    cboTermQuotes.AddItem "Double quotes"
    cboTermQuotes.ListIndex = 1

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

    Set btnRun = AddButton(Me, "btnRun", "Run Checks", PAD, yPos, ACT_BTN_W, ACT_BTN_H, , True)
    Set btnExport = AddButton(Me, "btnExport", "Export Report", PAD + ACT_BTN_W + ACT_GAP, yPos, ACT_BTN_W, ACT_BTN_H)
    Set btnClose = AddButton(Me, "btnClose", "Close", PAD + 2 * (ACT_BTN_W + ACT_GAP), yPos, 70, ACT_BTN_H)

    yPos = yPos + ACT_BTN_H + ITEM_GAP

    ' ==========================================================
    '  ROW 5: Status Bar
    ' ==========================================================
    Set lblStatus = AddLabel(Me, "lblStatus", "Ready. Select rules and click Run.", PAD, yPos, FULL_W, LBL_H, 8)

    ' -- Load custom rules from engine -------------------------
    LoadCustomRulesFromEngine

    ' -- Final form size based on layout ---
    Dim neededH As Single
    neededH = yPos + LBL_H + PAD
    If neededH < 300 Then neededH = 300

    Me.Width = FULL_W + 2 * PAD
    Me.Height = neededH

    Debug.Print "frmPleadingsChecker_Initialize: Width=" & Me.Width & " Height=" & Me.Height
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
'  Single source of truth: crCorrect(), crVariants(), crInsertSeq()
' ============================================================
Private Sub InitCustomRulesArrays(ByVal capacity As Long)
    If capacity < 1 Then capacity = 1
    ReDim crCorrect(0 To capacity - 1)
    ReDim crVariants(0 To capacity - 1)
    ReDim crInsertSeq(0 To capacity - 1)
    ReDim crSortOrder(0 To capacity - 1)
    crCount = 0
    crNextSeq = 1
End Sub

Private Sub AddCustomRule(ByVal correct As String, ByVal variants As String, _
                          Optional ByVal seq As Long = -1)
    If crCount = 0 Then
        InitCustomRulesArrays 16
    End If
    ' Grow arrays if needed
    If crCount > UBound(crCorrect) Then
        ReDim Preserve crCorrect(0 To crCount * 2)
        ReDim Preserve crVariants(0 To crCount * 2)
        ReDim Preserve crInsertSeq(0 To crCount * 2)
        ReDim Preserve crSortOrder(0 To crCount * 2)
    End If
    crCorrect(crCount) = correct
    crVariants(crCount) = variants
    If seq > 0 Then
        crInsertSeq(crCount) = seq
        If seq >= crNextSeq Then crNextSeq = seq + 1
    Else
        crInsertSeq(crCount) = crNextSeq
        crNextSeq = crNextSeq + 1
    End If
    crCount = crCount + 1
    RebuildSortOrder
End Sub

Private Sub RemoveCustomRule(ByVal idx As Long)
    If idx < 0 Or idx >= crCount Then Exit Sub
    Dim j As Long
    For j = idx To crCount - 2
        crCorrect(j) = crCorrect(j + 1)
        crVariants(j) = crVariants(j + 1)
        crInsertSeq(j) = crInsertSeq(j + 1)
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
    SortOrderByField crSortMode
    If Not crSortAscending Then ReverseSortOrder
End Sub

Private Sub ReverseSortOrder()
    Dim lo As Long, hi As Long, tmp As Long
    lo = 0: hi = crCount - 1
    Do While lo < hi
        tmp = crSortOrder(lo)
        crSortOrder(lo) = crSortOrder(hi)
        crSortOrder(hi) = tmp
        lo = lo + 1: hi = hi - 1
    Loop
End Sub

Private Sub SortOrderByField(ByVal field As Long)
    ' Insertion sort on crSortOrder by chosen field
    ' 0 = insertion sequence, 1 = correct, 2 = variants
    If crCount < 2 Then Exit Sub
    Dim i As Long, j As Long, tmp As Long
    Dim shouldStop As Boolean
    For i = 1 To crCount - 1
        tmp = crSortOrder(i)
        j = i - 1
        Do While j >= 0
            If field = 0 Then
                shouldStop = (crInsertSeq(crSortOrder(j)) <= crInsertSeq(tmp))
            ElseIf field = 1 Then
                shouldStop = (LCase$(crCorrect(crSortOrder(j))) <= LCase$(crCorrect(tmp)))
            Else
                shouldStop = (LCase$(crVariants(crSortOrder(j))) <= LCase$(crVariants(tmp)))
            End If
            If shouldStop Then Exit Do
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
'  LOAD CUSTOM RULES FROM ENGINE
' ============================================================
Private Sub LoadCustomRulesFromEngine()
    InitCustomRulesArrays 16
    On Error Resume Next
    Dim engineRules As Object
    Set engineRules = Application.Run("Rules_Brands.GetBrandRules")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        RefreshCustomRulesList
        Exit Sub
    End If
    On Error GoTo 0
    If engineRules Is Nothing Then
        RefreshCustomRulesList
        Exit Sub
    End If
    Dim rKey As Variant
    For Each rKey In engineRules.keys
        AddCustomRule CStr(rKey), CStr(engineRules(rKey))
    Next rKey
    RefreshCustomRulesList
End Sub

' ============================================================
'  SYNC CUSTOM RULES BACK TO ENGINE (before run)
' ============================================================
Private Sub SyncCustomRulesToEngine()
    ' Clear existing engine rules and rebuild from our data model
    On Error Resume Next
    Dim existingRules As Object
    Set existingRules = Application.Run("Rules_Brands.GetBrandRules")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    ' Remove all existing rules from engine
    If Not existingRules Is Nothing Then
        Dim oldKeys As Variant
        oldKeys = existingRules.keys
        Dim m As Long
        For m = UBound(oldKeys) To LBound(oldKeys) Step -1
            On Error Resume Next
            Application.Run "Rules_Brands.RemoveBrandRule", CStr(oldKeys(m))
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
        Next m
    End If

    ' Add all custom rules to engine
    Dim n As Long
    For n = 0 To crCount - 1
        On Error Resume Next
        Application.Run "Rules_Brands.AddBrandRule", crCorrect(n), crVariants(n)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
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
    For i = 0 To crCount - 1
        di = crSortOrder(i)
        lstCustomRules.AddItem ""
        lstCustomRules.List(i, 0) = CStr(crInsertSeq(di))
        lstCustomRules.List(i, 1) = crCorrect(di)
        lstCustomRules.List(i, 2) = crVariants(di)
    Next i
End Sub

' ============================================================
'  GATHER FORM CONFIG (Section G)
'  Collects all UI state into a single dictionary for the engine.
' ============================================================
Private Function GatherFormConfig() As Object
    Dim cfg As Object
    Set cfg = CreateObject("Scripting.Dictionary")

    ' Sync rule config from dynamic checkboxes
    Dim i As Long
    For i = 1 To ruleCheckboxes.Count
        Dim rName As String
        rName = ruleKeys(i - 1)
        If ruleConfig.Exists(rName) Then
            ruleConfig(rName) = CBool(ruleCheckboxes(i).Value)
        End If
    Next i
    Set cfg("ruleConfig") = ruleConfig

    ' Page range (ignore placeholder text)
    cfg("pageRange") = GetPageRangeInput()

    ' Spelling mode
    If cboSpelling.ListIndex = 1 Then
        cfg("spellingMode") = "US"
    Else
        cfg("spellingMode") = "UK"
    End If

    ' Quote nesting
    If cboQuoteNesting.ListIndex = 1 Then
        cfg("quoteNesting") = "DOUBLE"
    Else
        cfg("quoteNesting") = "SINGLE"
    End If

    ' Smart quotes
    If cboSmartQuotes.ListIndex = 1 Then
        cfg("smartQuotePref") = "STRAIGHT"
    Else
        cfg("smartQuotePref") = "SMART"
    End If

    ' Date format
    If cboDateFormat.ListIndex = 1 Then
        cfg("dateFormatPref") = "US"
    Else
        cfg("dateFormatPref") = "UK"
    End If

    ' Non-English terms
    If cboNonEngTerms.ListIndex = 1 Then
        cfg("nonEngTermPref") = "REGULAR"
    Else
        cfg("nonEngTermPref") = "ITALICS"
    End If

    ' Term format
    Dim termFmt As String
    Select Case cboTermFormat.ListIndex
        Case 0: termFmt = "BOLD"
        Case 1: termFmt = "BOLDITALIC"
        Case 2: termFmt = "ITALIC"
        Case Else: termFmt = "NONE"
    End Select
    cfg("termFormatPref") = termFmt

    ' Term quotes
    If cboTermQuotes.ListIndex = 0 Then
        cfg("termQuotePref") = "SINGLE"
    Else
        cfg("termQuotePref") = "DOUBLE"
    End If

    ' Space style
    If cboSpaceStyle.ListIndex = 1 Then
        cfg("spaceStylePref") = "TWO"
    Else
        cfg("spaceStylePref") = "ONE"
    End If

    Set GatherFormConfig = cfg
End Function

' ============================================================
'  RUN BUTTON
' ============================================================
Private Sub btnRun_Click()
    Set targetDoc = PleadingsEngine.GetTargetDocument()
    If targetDoc Is Nothing Then
        Exit Sub
    End If

    ' Sync custom rules to engine
    SyncCustomRulesToEngine

    ' Gather all UI state into a single config object
    Dim formCfg As Object
    Set formCfg = GatherFormConfig()

    ' Reset cancel flag before run
    PleadingsEngine.ResetCancelRun

    ' Run checks with cancellation support
    lblStatus.Caption = "Running checks..."
    Me.Repaint
    DoEvents

    On Error GoTo RunCancelled

    Set lastResults = PleadingsEngine.RunCheckerFromFormConfig(targetDoc, formCfg)

    ' Show performance summary in Immediate window
    Dim slowestRules As String
    If PleadingsEngine.ENABLE_PROFILING Then
        Dim perfSummary As String
        perfSummary = PleadingsEngine.GetPerformanceSummary()
        slowestRules = PleadingsEngine.GetTopSlowestRules(3)
        Debug.Print "frmPleadingsChecker final: Width=" & Me.Width & " Height=" & Me.Height
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

    Exit Sub

RunCancelled:
    If Err.Number = vbObjectError + 513 Then
        lblStatus.Caption = "Run cancelled."
        MsgBox "Run cancelled.", vbInformation, "Pleadings Checker"
    Else
        lblStatus.Caption = "Error: " & Err.Description
        MsgBox "An error occurred:" & vbCrLf & vbCrLf & Err.Description, _
               vbExclamation, "Pleadings Checker"
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

    ' Generate JSON report
    Dim reportSummary As String
    reportSummary = PleadingsEngine.GenerateReport(lastResults, reportPath, targetDoc)

    ' Generate plain-text report alongside JSON
    Dim txtPath As String
    If Len(reportPath) > 5 And LCase$(Right$(reportPath, 5)) = ".json" Then
        txtPath = Left$(reportPath, Len(reportPath) - 5) & ".txt"
    Else
        txtPath = reportPath & ".txt"
    End If

    Dim txtSummary As String
    On Error Resume Next
    txtSummary = PleadingsEngine.GenerateTextReport(lastResults, txtPath, targetDoc)
    If Err.Number <> 0 Then
        txtSummary = "Text report generation failed: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0

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
    msg = "Reports saved:" & vbCrLf & _
          "  JSON: " & reportPath & vbCrLf & _
          "  Text: " & txtPath

    If logSaved And Len(logPath) > 0 Then
        msg = msg & vbCrLf & vbCrLf & "Debug log: " & logPath
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
    correctForm = GetCorrectText()
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
        ' Add new rule
        AddCustomRule correctForm, incorrectVars
    End If

    ClearPlaceholderText txtRuleCorrect, mCorrectShowingPlaceholder
    InitPlaceholder txtRuleCorrect, CORRECT_PLACEHOLDER, mCorrectShowingPlaceholder
    ClearPlaceholderText txtRuleVariants, mVariantsShowingPlaceholder
    InitPlaceholder txtRuleVariants, VARIANTS_PLACEHOLDER, mVariantsShowingPlaceholder
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
        ClearPlaceholderText txtRuleCorrect, mCorrectShowingPlaceholder
        InitPlaceholder txtRuleCorrect, CORRECT_PLACEHOLDER, mCorrectShowingPlaceholder
        ClearPlaceholderText txtRuleVariants, mVariantsShowingPlaceholder
        InitPlaceholder txtRuleVariants, VARIANTS_PLACEHOLDER, mVariantsShowingPlaceholder
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
    ClearPlaceholderText txtRuleCorrect, mCorrectShowingPlaceholder
    txtRuleCorrect.Text = crCorrect(dataIdx)
    ClearPlaceholderText txtRuleVariants, mVariantsShowingPlaceholder
    txtRuleVariants.Text = crVariants(dataIdx)
    btnAddRule.Caption = "Save Edit"
End Sub

' ============================================================
'  CUSTOM RULES: HEADER-CLICK SORT (via clsHeaderClick class)
' ============================================================
Public Sub HandleHeaderSort(ByVal sortField As Long)
    If crSortMode = sortField Then
        ' Same header clicked again: reverse direction
        crSortAscending = Not crSortAscending
    Else
        crSortMode = sortField
        crSortAscending = True
    End If
    RebuildSortOrder
    RefreshCustomRulesList
End Sub

' ============================================================
'  CUSTOM RULES: SAVE
' ============================================================
Private Sub btnSaveRules_Click()
    If crCount = 0 Then
        MsgBox "No custom rules to save.", vbExclamation, "Custom Rules"
        Exit Sub
    End If

    Dim rulesFile As String
    rulesFile = GetCustomRulesPath()

    Dim rulesDir As String
    rulesDir = GetParentDirectory(rulesFile)
    If Len(rulesDir) > 0 Then
        If Not EnsureDirectoryExists(rulesDir) Then
            MsgBox "Could not create directory:" & vbCrLf & rulesDir & vbCrLf & vbCrLf & _
                   "Check permissions and try again.", vbExclamation, "Custom Rules"
            Exit Sub
        End If
    End If

    ' Save all rules as tab-delimited: seq<TAB>correct<TAB>variants
    Dim fileNum As Integer
    fileNum = FreeFile
    On Error GoTo SaveFail
    Open rulesFile For Output As #fileNum
    Print #fileNum, "# Custom Rules (tab-delimited: seq, correct, variants)"
    Dim s As Long
    For s = 0 To crCount - 1
        Print #fileNum, CStr(crInsertSeq(s)) & vbTab & crCorrect(s) & vbTab & crVariants(s)
    Next s
    Close #fileNum

    MsgBox "Custom rules saved (" & crCount & " rules) to:" & vbCrLf & rulesFile, _
           vbInformation, "Pleadings Checker"
    Exit Sub

SaveFail:
    Dim saveErrNum As Long, saveErrDesc As String
    saveErrNum = Err.Number
    saveErrDesc = Err.Description
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    MsgBox "Failed to save custom rules." & vbCrLf & vbCrLf & _
           "File: " & rulesFile & vbCrLf & _
           "Error " & saveErrNum & ": " & saveErrDesc, vbExclamation, "Pleadings Checker"
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
    Dim correct As String
    Dim variants As String
    Dim seq As Long
    Dim parts() As String

    InitCustomRulesArrays 16

    fileNum = FreeFile
    On Error GoTo LoadFail
    Open filePath For Input As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim$(lineText)
        If Len(lineText) = 0 Then GoTo NextLoadLine
        If Left$(lineText, 1) = "#" Then GoTo NextLoadLine

        ' Strip legacy +/- prefix if present
        If Left$(lineText, 1) = "+" Or Left$(lineText, 1) = "-" Then
            lineText = Mid$(lineText, 2)
        End If

        ' Try tab-delimited: seq<TAB>correct<TAB>variants
        If InStr(lineText, vbTab) > 0 Then
            parts = Split(lineText, vbTab)
            If UBound(parts) >= 2 Then
                seq = 0
                On Error Resume Next
                seq = CLng(parts(0))
                If Err.Number <> 0 Then seq = 0: Err.Clear
                On Error GoTo LoadFail
                correct = Trim$(parts(1))
                variants = Trim$(parts(2))
                If Len(correct) > 0 And Len(variants) > 0 Then
                    If seq > 0 Then
                        AddCustomRule correct, variants, seq
                    Else
                        AddCustomRule correct, variants
                    End If
                End If
            End If
        Else
            ' Legacy fallback: correct=variants
            Dim eqPos As Long
            eqPos = InStr(lineText, "=")
            If eqPos > 1 Then
                correct = Trim$(Left$(lineText, eqPos - 1))
                variants = Trim$(Mid$(lineText, eqPos + 1))
                If Len(correct) > 0 And Len(variants) > 0 Then
                    AddCustomRule correct, variants
                End If
            End If
        End If

NextLoadLine:
    Loop

    Close #fileNum
    RefreshCustomRulesList
    SyncCustomRulesToEngine

    MsgBox "Custom rules loaded from:" & vbCrLf & filePath, vbInformation, "Custom Rules"
    Exit Sub

LoadFail:
    On Error Resume Next
    Close #fileNum
    MsgBox "Could not read custom rules from:" & vbCrLf & filePath & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Custom Rules"
    Err.Clear
    On Error GoTo 0
    If crCount = 0 Then LoadCustomRulesFromEngine
End Sub

' -- Helper: cross-platform custom rules file path --
Private Function GetCustomRulesPath() As String
    Dim sep As String
    sep = Application.PathSeparator
    #If Mac Then
        GetCustomRulesPath = Environ("HOME") & sep & "Library" & sep & _
                            "Application Support" & sep & "PleadingsChecker" & sep & "custom_rules.txt"
    #Else
        GetCustomRulesPath = Environ("APPDATA") & sep & "PleadingsChecker" & sep & "custom_rules.txt"
    #End If
End Function

' ============================================================
'  SHARED PLACEHOLDER HELPERS
'  Generic init / get / clear / enter / exit for any textbox.
' ============================================================
Private Sub InitPlaceholder(ByVal tb As MSForms.TextBox, _
                            ByVal placeholder As String, _
                            ByRef flag As Boolean)
    If tb Is Nothing Then Exit Sub
    tb.Text = placeholder
    tb.ForeColor = RGB(150, 150, 150)
    flag = True
End Sub

Private Function GetPlaceholderText(ByVal tb As MSForms.TextBox, _
                                    ByVal placeholder As String, _
                                    ByVal showing As Boolean) As String
    If showing Then
        GetPlaceholderText = vbNullString
        Exit Function
    End If
    Dim raw As String
    raw = Trim$(tb.Text)
    If raw = placeholder Then
        GetPlaceholderText = vbNullString
    Else
        GetPlaceholderText = raw
    End If
End Function

Private Sub ClearPlaceholderText(ByVal tb As MSForms.TextBox, ByRef flag As Boolean)
    tb.Text = ""
    tb.ForeColor = RGB(0, 0, 0)
    flag = False
End Sub

Private Sub HandlePlaceholderEnter(ByVal tb As MSForms.TextBox, ByRef flag As Boolean)
    If flag Then
        tb.Text = ""
        tb.ForeColor = RGB(0, 0, 0)
        flag = False
    End If
End Sub

Private Sub HandlePlaceholderExit(ByVal tb As MSForms.TextBox, _
                                  ByVal placeholder As String, _
                                  ByRef flag As Boolean)
    If Len(Trim$(tb.Text)) = 0 Then
        InitPlaceholder tb, placeholder, flag
    Else
        tb.ForeColor = RGB(0, 0, 0)
        flag = False
    End If
End Sub

' -- Thin wrappers for the three placeholdered textboxes (keep named accessors) --
Private Function GetCorrectText() As String
    GetCorrectText = GetPlaceholderText(txtRuleCorrect, CORRECT_PLACEHOLDER, mCorrectShowingPlaceholder)
End Function

Private Function GetVariantsText() As String
    GetVariantsText = GetPlaceholderText(txtRuleVariants, VARIANTS_PLACEHOLDER, mVariantsShowingPlaceholder)
End Function

' Event handlers must be separate subs; they delegate to shared helpers.
Private Sub txtRuleCorrect_Enter()
    HandlePlaceholderEnter txtRuleCorrect, mCorrectShowingPlaceholder
End Sub

Private Sub txtRuleCorrect_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    HandlePlaceholderExit txtRuleCorrect, CORRECT_PLACEHOLDER, mCorrectShowingPlaceholder
End Sub

Private Sub txtRuleVariants_Enter()
    HandlePlaceholderEnter txtRuleVariants, mVariantsShowingPlaceholder
End Sub

Private Sub txtRuleVariants_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    HandlePlaceholderExit txtRuleVariants, VARIANTS_PLACEHOLDER, mVariantsShowingPlaceholder
End Sub

' Normalise comma-separated variants: trim each, remove blanks
Private Function NormaliseVariants(ByVal raw As String) As String
    Dim parts() As String
    parts = Split(raw, ",")
    Dim result As String
    Dim p As Long
    For p = LBound(parts) To UBound(parts)
        Dim partStr As String
        partStr = Trim$(parts(p))
        If Len(partStr) > 0 Then
            If Len(result) > 0 Then result = result & ", "
            result = result & partStr
        End If
    Next p
    NormaliseVariants = result
End Function

' -- Page-range placeholder (delegates to shared helpers) --
Public Function GetPageRangeInput() As String
    GetPageRangeInput = GetPlaceholderText(txtPageRange, PAGE_RANGE_PLACEHOLDER, mPageRangeShowingPlaceholder)
End Function

Private Sub txtPageRange_Enter()
    HandlePlaceholderEnter txtPageRange, mPageRangeShowingPlaceholder
End Sub

Private Sub txtPageRange_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    HandlePlaceholderExit txtPageRange, PAGE_RANGE_PLACEHOLDER, mPageRangeShowingPlaceholder
End Sub

' ============================================================
'  CLOSE BUTTON
' ============================================================
Private Sub btnClose_Click()
    Unload Me
End Sub
