VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHeaderClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' clsHeaderClick.cls
' Event-sink class to handle Click on runtime MSForms.Label
' controls used as sortable column headers in the custom-rules
' table of frmPleadingsChecker.
'
' Usage from the form:
'   Dim mHdr As New clsHeaderClick
'   mHdr.Init Me, someLabelCtrl, 1
'
' When the label is clicked, the class calls back into the
' form's Public HandleHeaderSort method.
' ============================================================
Option Explicit

Public WithEvents LabelCtrl As MSForms.Label

Private mCallback  As Object   ' late-bound ref to frmPleadingsChecker
Private mSortField As Long     ' 0 = default/#, 1 = correct, 2 = variants

Public Sub Init(ByVal parentForm As Object, _
                ByVal lbl As MSForms.Label, _
                ByVal sortField As Long)
    Set mCallback = parentForm
    Set LabelCtrl = lbl
    mSortField = sortField
End Sub

Private Sub LabelCtrl_Click()
    If Not mCallback Is Nothing Then
        mCallback.HandleHeaderSort mSortField
    End If
End Sub

Private Sub Class_Terminate()
    Set LabelCtrl = Nothing
    Set mCallback = Nothing
End Sub
