VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA006002F3} frmWordChecker
   Caption         =   "Word Checker"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   8040
   StartUpPosition =   1  'CenterOwner
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnToggleTrack
      Caption         =   "Track Changes: OFF"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   2040
   End
   ' ── Tab toggle buttons ──────────────────────────────────
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnTabConventions
      Caption         =   "1. Conventions"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1800
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnTabChecks
      Caption         =   "2. Run Checks"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1800
   End
   Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnTabActions
      Caption         =   "3. Actions"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   1800
   End
   ' ── FRAME 1: Conventions ────────────────────────────────
   Begin {6E182020-7460-11CE-9E0D-00AA006002F3} fraConventions
      Caption         =   "Formatting Conventions"
      Height          =   9000
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   7800
      ' Case names
      Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} Label1
         Caption         =   "Case names:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1320
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optCaseUnderline
         Caption         =   "Underline"
         Height          =   255
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1200
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optCaseItalic
         Caption         =   "Italic"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   960
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optCaseBoth
         Caption         =   "Both"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   240
         Width           =   960
      End
      ' v. / p. full stops
      Begin {8BD21D40-EC42-11CE-9E0D-00AA006002F3} chkVDot
         Caption         =   "Use ""v."" (full stop) in case names"
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   600
         Width           =   3360
      End
      Begin {8BD21D40-EC42-11CE-9E0D-00AA006002F3} chkPDot
         Caption         =   "Use ""p."" (full stop) for page references"
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   960
         Width           =   3600
      End
      ' Spacing after full stop
      Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} Label2
         Caption         =   "Spacing after full stop:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   2040
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optSpaceSingle
         Caption         =   "Single space"
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   1320
         Value           =   -1  'True
         Width           =   1440
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optSpaceDouble
         Caption         =   "Double space"
         Height          =   255
         Left            =   3840
         TabIndex        =   18
         Top             =   1320
         Width           =   1440
      End
      ' i.e.
      Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} Label3
         Caption         =   "i.e. write as:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   1320
      End
      Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} txtIeFormat
         Height          =   315
         Left            =   1560
         TabIndex        =   20
         Text            =   "i.e."
         Top             =   1680
         Width           =   960
      End
      Begin {8BD21D40-EC42-11CE-9E0D-00AA006002F3} chkIeComma
         Caption         =   "Comma after (i.e.,)"
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   1680
         Width           =   2160
      End
      ' e.g.
      Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} Label4
         Caption         =   "e.g. write as:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2040
         Width           =   1320
      End
      Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} txtEgFormat
         Height          =   315
         Left            =   1560
         TabIndex        =   23
         Text            =   "e.g."
         Top             =   2040
         Width           =   960
      End
      Begin {8BD21D40-EC42-11CE-9E0D-00AA006002F3} chkEgComma
         Caption         =   "Comma after (e.g.,)"
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   2040
         Width           =   2160
      End
      ' Edition
      Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} Label5
         Caption         =   "Edition (fn only):"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   1440
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optEdEd
         Caption         =   "ed"
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   2400
         Value           =   -1  'True
         Width           =   720
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optEdEdn
         Caption         =   "edn"
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   2400
         Width           =   840
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optEdEdition
         Caption         =   "edition"
         Height          =   255
         Left            =   3480
         TabIndex        =   28
         Top             =   2400
         Width           =   960
      End
      ' Emphasis
      Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} Label6
         Caption         =   "Emphasis style:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2760
         Width           =   1440
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optEmphUnderline
         Caption         =   "Underline"
         Height          =   255
         Left            =   1680
         TabIndex        =   30
         Top             =   2760
         Value           =   -1  'True
         Width           =   1200
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optEmphBold
         Caption         =   "Bold"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   2760
         Width           =   840
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optEmphItalic
         Caption         =   "Italic"
         Height          =   255
         Left            =   3960
         TabIndex        =   32
         Top             =   2760
         Width           =   840
      End
      ' Block quotes
      Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} Label7
         Caption         =   "Block quotes:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3120
         Width           =   1320
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optBlockSmaller
         Caption         =   "Smaller font, no italics"
         Height          =   255
         Left            =   1560
         TabIndex        =   34
         Top             =   3120
         Value           =   -1  'True
         Width           =   2280
      End
      Begin {8BD21D50-EC42-11CE-9E0D-00AA006002F3} optBlockItalic
         Caption         =   "Size 12 + italics"
         Height          =   255
         Left            =   3960
         TabIndex        =   35
         Top             =   3120
         Width           =   1920
      End
      ' Cross-references
      Begin {8BD21D40-EC42-11CE-9E0D-00AA006002F3} chkNoSupraInfra
         Caption         =   "Flag supra / infra / ibid (no internal cross-references)"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3480
         Value           =   -1  'True (checked by default)
         Width           =   5040
      End
      ' Citation format
      Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} Label8
         Caption         =   "Citation template:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3840
         Width           =   1560
      End
      Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} txtCitationFormat
         Height          =   315
         Left            =   1800
         TabIndex        =   38
         Text            =   "[Exhibit], DD Month YYYY (Short Name), Exhibit R/C-[ ]"
         Top             =   3840
         Width           =   5760
      End
      ' Hint label
      Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblConvHint
         Caption         =   "Set your conventions here before running checks. Settings are saved with the document."
         ForeColor       =   &H00808080&
         Height          =   360
         Left            =   120
         TabIndex        =   39
         Top             =   4260
         Width           =   7440
      End
   End
   ' ── FRAME 2: Run Checks ─────────────────────────────────
   Begin {6E182020-7460-11CE-9E0D-00AA006002F3} fraChecks
      Caption         =   "Run Checks"
      Height          =   9000
      Left            =   120
      TabIndex        =   40
      Top             =   600
      Visible         =   0  'False
      Width           =   7800
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckLegalBlobs
         Caption         =   "Check Blobs (unfilled placeholders)"
         BackColor       =   &H000000FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckCaseNames
         Caption         =   "Check Case Names"
         Height          =   390
         Left            =   120
         TabIndex        =   51
         Top             =   720
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckSpacing
         Caption         =   "Check Spacing After Full Stop"
         Height          =   390
         Left            =   120
         TabIndex        =   52
         Top             =   1200
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckIeEg
         Caption         =   "Check i.e. / e.g."
         Height          =   390
         Left            =   120
         TabIndex        =   53
         Top             =   1680
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckEdition
         Caption         =   "Check Edition Abbrev. (footnotes)"
         Height          =   390
         Left            =   120
         TabIndex        =   54
         Top             =   2160
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckEllipses
         Caption         =   "Check Ellipses Spacing"
         Height          =   390
         Left            =   120
         TabIndex        =   55
         Top             =   2640
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckPinpoints
         Caption         =   "Check Pinpoints (no ""at"")"
         Height          =   390
         Left            =   120
         TabIndex        =   56
         Top             =   3120
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckCrossRefs
         Caption         =   "Check Cross-References (supra/infra)"
         Height          =   390
         Left            =   4200
         TabIndex        =   57
         Top             =   240
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckCitations
         Caption         =   "Check Citations / Exhibits"
         Height          =   390
         Left            =   4200
         TabIndex        =   58
         Top             =   720
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckBrackets
         Caption         =   "Check Brackets in Quotes (no italic)"
         Height          =   390
         Left            =   4200
         TabIndex        =   59
         Top             =   1200
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckTables
         Caption         =   "Check Table Spacing (6pt)"
         Height          =   390
         Left            =   4200
         TabIndex        =   60
         Top             =   1680
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckCapitalisation
         Caption         =   "Check Capitalisation (clause/Clause)"
         Height          =   390
         Left            =   4200
         TabIndex        =   61
         Top             =   2160
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnCheckDashes
         Caption         =   "Check Dashes (en/em/hyphen)"
         Height          =   390
         Left            =   4200
         TabIndex        =   62
         Top             =   2640
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnRunAll
         Caption         =   "▶  RUN ALL CHECKS"
         Default         =   -1  'True
         Height          =   510
         Left            =   120
         TabIndex        =   63
         Top             =   3720
         Width           =   7560
      End
   End
   ' ── FRAME 3: Actions ────────────────────────────────────
   Begin {6E182020-7460-11CE-9E0D-00AA006002F3} fraActions
      Caption         =   "Document Actions"
      Height          =   9000
      Left            =   120
      TabIndex        =   70
      Top             =   600
      Visible         =   0  'False
      Width           =   7800
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnUpdateTOC
         Caption         =   "Update Table of Contents"
         Height          =   510
         Left            =   120
         TabIndex        =   71
         Top             =   240
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnUpdateCrossRefs
         Caption         =   "Update Cross-References (F9)"
         Height          =   510
         Left            =   120
         TabIndex        =   72
         Top             =   840
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnSendToDS
         Caption         =   "Draft Email to Document Services"
         Height          =   510
         Left            =   4200
         TabIndex        =   73
         Top             =   240
         Width           =   3480
      End
      Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblDSHint
         Caption         =   "Drafts email to Global Document Specialists at Freshfields (requires Outlook)"
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   4200
         TabIndex        =   74
         Top             =   840
         Width           =   3480
      End
      Begin {D7053240-CE69-11CD-A777-00DD01143C57} btnClose
         Cancel          =   -1  'True
         Caption         =   "Close"
         Height          =   510
         Left            =   120
         TabIndex        =   75
         Top             =   1440
         Width           =   1680
      End
   End
   ' ── Results pane (always visible) ───────────────────────
   Begin {978C9E23-D4B0-11CE-BF2D-00AA003F40D0} lblResults
      Caption         =   "Results:"
      Height          =   255
      Left            =   120
      TabIndex        =   90
      Top             =   9720
      Width           =   840
   End
   Begin {8BD21D10-EC42-11CE-9E0D-00AA006002F3} txtResults
      Height          =   720
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   91
      Top             =   9960
      Width           =   7800
   End
End
Attribute VB_Name = "frmWordChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ── On load ──────────────────────────────────────────────────
' ── Guard: document still open? ─────────────────────────────
' Returns True and does nothing if a document is active.
' Returns False and shows a message if all documents have been closed
' while the form was still open (the entry-point sub already checks
' for a document before showing the form, so this only fires if the
' user closes the document mid-session).
Private Function DocIsReady() As Boolean
    If ActiveDocument Is Nothing Then
        MsgBox "No document is open.", vbExclamation, "Word Checker"
        DocIsReady = False
    Else
        DocIsReady = True
    End If
End Function

' ── On load ──────────────────────────────────────────────────
Private Sub UserForm_Initialize()
    ' Default to Conventions tab visible
    ShowFrame "conventions"
    UpdateTrackCaption

    ' Font.Bold cannot be set in the .frm design section for CommandButton;
    ' set it here instead.
    btnRunAll.Font.Bold = True

    ' Load any saved conventions from document properties
    LoadConventions
End Sub

' ── Tab switching ────────────────────────────────────────────
Private Sub btnTabConventions_Click()
    ShowFrame "conventions"
End Sub

Private Sub btnTabChecks_Click()
    ShowFrame "checks"
End Sub

Private Sub btnTabActions_Click()
    ShowFrame "actions"
End Sub

Private Sub ShowFrame(which As String)
    fraConventions.Visible = (which = "conventions")
    fraChecks.Visible      = (which = "checks")
    fraActions.Visible     = (which = "actions")

    ' Bold the active tab button caption as visual indicator
    btnTabConventions.Font.Bold = (which = "conventions")
    btnTabChecks.Font.Bold      = (which = "checks")
    btnTabActions.Font.Bold     = (which = "actions")
End Sub

' ── Tracked changes toggle ───────────────────────────────────
Private Sub btnToggleTrack_Click()
    If Not DocIsReady() Then Exit Sub
    ToggleTrackedChanges ActiveDocument
    UpdateTrackCaption
End Sub

Private Sub UpdateTrackCaption()
    If ActiveDocument Is Nothing Then
        btnToggleTrack.Caption = "Track Changes: —"
        btnToggleTrack.BackColor = &H8000000F
        Exit Sub
    End If
    btnToggleTrack.Caption = TrackChangesCaption(ActiveDocument)
    If ActiveDocument.TrackRevisions Then
        btnToggleTrack.BackColor = RGB(255, 200, 200)  ' light red when ON
    Else
        btnToggleTrack.BackColor = &H8000000F  ' default button colour
    End If
End Sub

' ── Helper: read settings from form ─────────────────────────
Private Function GetCaseStyle() As String
    If optCaseUnderline.Value Then
        GetCaseStyle = "underline"
    ElseIf optCaseItalic.Value Then
        GetCaseStyle = "italic"
    Else
        GetCaseStyle = "both"
    End If
End Function

Private Function GetEdStyle() As String
    If optEdEd.Value Then
        GetEdStyle = "ed"
    ElseIf optEdEdn.Value Then
        GetEdStyle = "edn"
    Else
        GetEdStyle = "edition"
    End If
End Function

' ── Log a result ────────────────────────────────────────────
Private Sub LogResult(msg As String)
    Dim existing As String
    existing = txtResults.Text
    If Len(existing) > 5000 Then existing = "[earlier results cleared]"
    If existing <> "" Then existing = existing & vbCrLf
    txtResults.Text = existing & msg
    ' Scroll to bottom
    txtResults.SelStart = Len(txtResults.Text)
End Sub

' ── Individual check buttons ─────────────────────────────────
Private Sub btnCheckLegalBlobs_Click()
    If Not DocIsReady() Then Exit Sub
    SaveConventions
    LogResult CheckLegalBlobs(ActiveDocument)
End Sub

Private Sub btnCheckCaseNames_Click()
    If Not DocIsReady() Then Exit Sub
    SaveConventions
    LogResult CheckCaseNames(ActiveDocument, GetCaseStyle(), chkVDot.Value)
End Sub

Private Sub btnCheckSpacing_Click()
    If Not DocIsReady() Then Exit Sub
    SaveConventions
    LogResult CheckSpacing(ActiveDocument, optSpaceDouble.Value)
End Sub

Private Sub btnCheckIeEg_Click()
    If Not DocIsReady() Then Exit Sub
    SaveConventions
    LogResult CheckIeEg(ActiveDocument, _
                        txtIeFormat.Text, txtEgFormat.Text, _
                        CBool(chkIeComma.Value), CBool(chkEgComma.Value))
End Sub

Private Sub btnCheckEdition_Click()
    If Not DocIsReady() Then Exit Sub
    SaveConventions
    LogResult CheckEdition(ActiveDocument, GetEdStyle())
End Sub

Private Sub btnCheckEllipses_Click()
    If Not DocIsReady() Then Exit Sub
    LogResult CheckEllipses(ActiveDocument)
End Sub

Private Sub btnCheckPinpoints_Click()
    If Not DocIsReady() Then Exit Sub
    LogResult CheckPinpoints(ActiveDocument)
End Sub

Private Sub btnCheckCrossRefs_Click()
    If Not DocIsReady() Then Exit Sub
    SaveConventions
    LogResult CheckCrossRefs(ActiveDocument, CBool(chkNoSupraInfra.Value))
End Sub

Private Sub btnCheckCitations_Click()
    If Not DocIsReady() Then Exit Sub
    SaveConventions
    LogResult CheckCitations(ActiveDocument, txtCitationFormat.Text)
End Sub

Private Sub btnCheckBrackets_Click()
    If Not DocIsReady() Then Exit Sub
    LogResult CheckBracketsInQuotes(ActiveDocument)
End Sub

Private Sub btnCheckTables_Click()
    If Not DocIsReady() Then Exit Sub
    LogResult CheckTables(ActiveDocument)
End Sub

Private Sub btnCheckCapitalisation_Click()
    If Not DocIsReady() Then Exit Sub
    LogResult CheckCapitalisation(ActiveDocument)
End Sub

Private Sub btnCheckDashes_Click()
    If Not DocIsReady() Then Exit Sub
    LogResult CheckDashes(ActiveDocument)
End Sub

' ── Run All ─────────────────────────────────────────────────
Private Sub btnRunAll_Click()
    If Not DocIsReady() Then Exit Sub
    SaveConventions
    txtResults.Text = "Running all checks..." & vbCrLf
    Me.Repaint

    Dim results As String
    results = RunAllChecks( _
        doc:=ActiveDocument, _
        caseStyle:=GetCaseStyle(), _
        vDot:=CBool(chkVDot.Value), _
        doubleSpace:=optSpaceDouble.Value, _
        ieConvention:=txtIeFormat.Text, _
        egConvention:=txtEgFormat.Text, _
        ieComma:=CBool(chkIeComma.Value), _
        egComma:=CBool(chkEgComma.Value), _
        edStyle:=GetEdStyle(), _
        noSupraInfra:=CBool(chkNoSupraInfra.Value), _
        citFormat:=txtCitationFormat.Text)

    txtResults.Text = results
    txtResults.SelStart = 0  ' scroll to top so blobs warning is visible first
End Sub

' ── Actions ─────────────────────────────────────────────────
Private Sub btnUpdateTOC_Click()
    If Not DocIsReady() Then Exit Sub
    UpdateTOC ActiveDocument
    LogResult "TOC updated."
End Sub

Private Sub btnUpdateCrossRefs_Click()
    If Not DocIsReady() Then Exit Sub
    UpdateCrossRefs ActiveDocument
    LogResult "Cross-reference fields updated."
End Sub

Private Sub btnSendToDS_Click()
    If Not DocIsReady() Then Exit Sub
    DraftDSEmail ActiveDocument
End Sub

Private Sub btnClose_Click()
    SaveConventions
    Unload Me
End Sub

' ── Persist conventions in document custom properties ────────
Private Sub SaveConventions()
    On Error Resume Next
    Dim doc As Document
    Set doc = ActiveDocument

    SetDocProp doc, "WC_CaseStyle",       GetCaseStyle()
    SetDocProp doc, "WC_VDot",            CStr(CBool(chkVDot.Value))
    SetDocProp doc, "WC_PDot",            CStr(CBool(chkPDot.Value))
    SetDocProp doc, "WC_DoubleSpace",     CStr(optSpaceDouble.Value)
    SetDocProp doc, "WC_IeFormat",        txtIeFormat.Text
    SetDocProp doc, "WC_EgFormat",        txtEgFormat.Text
    SetDocProp doc, "WC_IeComma",         CStr(CBool(chkIeComma.Value))
    SetDocProp doc, "WC_EgComma",         CStr(CBool(chkEgComma.Value))
    SetDocProp doc, "WC_EdStyle",         GetEdStyle()
    SetDocProp doc, "WC_NoSupraInfra",    CStr(CBool(chkNoSupraInfra.Value))
    SetDocProp doc, "WC_CitFormat",       txtCitationFormat.Text
    On Error GoTo 0
End Sub

Private Sub LoadConventions()
    On Error Resume Next
    Dim doc As Document
    Set doc = ActiveDocument

    Dim v As String

    v = GetDocProp(doc, "WC_CaseStyle")
    If v = "italic" Then
        optCaseItalic.Value = True
    ElseIf v = "both" Then
        optCaseBoth.Value = True
    ElseIf v = "underline" Then
        optCaseUnderline.Value = True
    End If

    v = GetDocProp(doc, "WC_VDot")
    If v <> "" Then chkVDot.Value = CBool(v)

    v = GetDocProp(doc, "WC_PDot")
    If v <> "" Then chkPDot.Value = CBool(v)

    v = GetDocProp(doc, "WC_DoubleSpace")
    If v = "True" Then
        optSpaceDouble.Value = True
    ElseIf v = "False" Then
        optSpaceSingle.Value = True
    End If

    v = GetDocProp(doc, "WC_IeFormat")
    If v <> "" Then txtIeFormat.Text = v

    v = GetDocProp(doc, "WC_EgFormat")
    If v <> "" Then txtEgFormat.Text = v

    v = GetDocProp(doc, "WC_IeComma")
    If v <> "" Then chkIeComma.Value = CBool(v)

    v = GetDocProp(doc, "WC_EgComma")
    If v <> "" Then chkEgComma.Value = CBool(v)

    v = GetDocProp(doc, "WC_EdStyle")
    If v = "edn" Then
        optEdEdn.Value = True
    ElseIf v = "edition" Then
        optEdEdition.Value = True
    ElseIf v = "ed" Then
        optEdEd.Value = True
    End If

    v = GetDocProp(doc, "WC_NoSupraInfra")
    If v <> "" Then chkNoSupraInfra.Value = CBool(v)

    v = GetDocProp(doc, "WC_CitFormat")
    If v <> "" Then txtCitationFormat.Text = v

    On Error GoTo 0
End Sub

' ── Document property helpers ────────────────────────────────
Private Sub SetDocProp(doc As Document, propName As String, propValue As String)
    On Error Resume Next
    doc.CustomDocumentProperties(propName).Value = propValue
    If Err.Number <> 0 Then
        Err.Clear
        doc.CustomDocumentProperties.Add _
            Name:=propName, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            Value:=propValue
    End If
    On Error GoTo 0
End Sub

Private Function GetDocProp(doc As Document, propName As String) As String
    On Error Resume Next
    GetDocProp = CStr(doc.CustomDocumentProperties(propName).Value)
    If Err.Number <> 0 Then GetDocProp = ""
    On Error GoTo 0
End Function
