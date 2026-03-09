Attribute VB_Name = "Rule22_BrandNameEnforcement"
' ============================================================
' Rule22_BrandNameEnforcement.bas
' Proofreading rule: enforces correct brand/entity name
' spellings and capitalisations. Maintains a configurable
' dictionary of correct forms and their known incorrect
' variants, and flags any incorrect usage found in the
' document.
'
' Provides persistence via SaveBrandRules / LoadBrandRules
' for user-customised brand lists.
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
'   - Microsoft Scripting Runtime (Scripting.Dictionary)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "brand_name_enforcement"

' ── Module-level brand rules dictionary ─────────────────────
' Key = correct form (String), Value = comma-separated incorrect variants (String)
Private brandRules As Scripting.Dictionary

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_BrandNameEnforcement(doc As Document) As Collection
    Dim issues As New Collection

    ' Initialise defaults if not yet loaded
    If brandRules Is Nothing Then
        InitDefaultBrands
    End If

    Dim keys As Variant
    Dim k As Long
    Dim correctForm As String
    Dim variants As Variant
    Dim v As Long
    Dim variant As String

    keys = brandRules.keys

    For k = 0 To brandRules.Count - 1
        correctForm = CStr(keys(k))
        variants = Split(CStr(brandRules(correctForm)), ",")

        For v = LBound(variants) To UBound(variants)
            variant = Trim(CStr(variants(v)))
            If Len(variant) = 0 Then GoTo NextVariant

            SearchAndFlag doc, variant, correctForm, issues

NextVariant:
        Next v
    Next k

    Set Check_BrandNameEnforcement = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Search for an incorrect variant and flag matches
' ════════════════════════════════════════════════════════════
Private Sub SearchAndFlag(doc As Document, _
                           variant As String, _
                           correctForm As String, _
                           ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim issue As PleadingsIssue
    Dim locStr As String

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = variant
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

        If PleadingsEngine.IsInPageRange(rng) Then
            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       "Incorrect brand name: '" & rng.Text & "'", _
                       "Use '" & correctForm & "'", _
                       rng.Start, _
                       rng.End, _
                       "error"
            issues.Add issue
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Populate default brand rules
' ════════════════════════════════════════════════════════════
Private Sub InitDefaultBrands()
    Set brandRules = New Scripting.Dictionary

    brandRules.Add "PwC", "PWC,Pwc,pwc"
    brandRules.Add "Deloitte", "deloitte,DELOITTE"
    brandRules.Add "HMRC", "Hmrc,hmrc,H.M.R.C."
    brandRules.Add "FCA", "Fca,fca,F.C.A."
    brandRules.Add "EY", "ey,Ernst & Young,Ernst and Young"
    brandRules.Add "KPMG", "kpmg,Kpmg"
End Sub

' ════════════════════════════════════════════════════════════
'  PUBLIC: Add or update a brand rule
' ════════════════════════════════════════════════════════════
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

' ════════════════════════════════════════════════════════════
'  PUBLIC: Remove a brand rule
' ════════════════════════════════════════════════════════════
Public Sub RemoveBrandRule(correct As String)
    If brandRules Is Nothing Then Exit Sub

    If brandRules.Exists(correct) Then
        brandRules.Remove correct
    End If
End Sub

' ════════════════════════════════════════════════════════════
'  PUBLIC: Get current brand rules dictionary
' ════════════════════════════════════════════════════════════
Public Function GetBrandRules() As Scripting.Dictionary
    If brandRules Is Nothing Then
        InitDefaultBrands
    End If

    Set GetBrandRules = brandRules
End Function

' ════════════════════════════════════════════════════════════
'  PUBLIC: Save brand rules to a text file
'  Format: one line per rule — "CorrectForm=variant1,variant2"
' ════════════════════════════════════════════════════════════
Public Sub SaveBrandRules(filePath As String)
    If brandRules Is Nothing Then Exit Sub

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
    Exit Sub

SaveError:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
End Sub

' ════════════════════════════════════════════════════════════
'  PUBLIC: Load brand rules from a text file
'  Replaces existing rules with contents of the file.
'  Format: one line per rule — "CorrectForm=variant1,variant2"
' ════════════════════════════════════════════════════════════
Public Sub LoadBrandRules(filePath As String)
    Dim fileNum As Integer
    Dim lineText As String
    Dim eqPos As Long
    Dim correct As String
    Dim variants As String

    Set brandRules = New Scripting.Dictionary

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
            variants = Trim(Mid(lineText, eqPos + 1))

            If Len(correct) > 0 And Len(variants) > 0 Then
                If Not brandRules.Exists(correct) Then
                    brandRules.Add correct, variants
                End If
            End If
        End If

NextLine:
    Loop

    Close #fileNum
    Exit Sub

LoadError:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    ' If file could not be loaded, fall back to defaults
    If brandRules Is Nothing Then
        InitDefaultBrands
    ElseIf brandRules.Count = 0 Then
        InitDefaultBrands
    End If
End Sub

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunBrandNameEnforcement()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Brand Name Enforcement"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_BrandNameEnforcement(doc)

    ' ── Highlight issues in document ─────────────────────────
    Dim iss As PleadingsIssue
    Dim rng As Range
    Dim i As Long
    For i = 1 To issues.Count
        Set iss = issues(i)
        If iss.RangeStart >= 0 And iss.RangeEnd > iss.RangeStart Then
            On Error Resume Next
            Set rng = doc.Range(iss.RangeStart, iss.RangeEnd)
            rng.HighlightColorIndex = wdYellow
            doc.Comments.Add Range:=rng, _
                Text:="[" & iss.RuleName & "] " & iss.Issue & _
                      " " & Chr(8212) & " Suggestion: " & iss.Suggestion
            On Error GoTo 0
        End If
    Next i

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Brand Name Enforcement"
End Sub
