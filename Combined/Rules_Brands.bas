Attribute VB_Name = "Rules_Brands"
' ============================================================
' Rules_Brands.bas
' Combined module for Rule 22: Brand Name Enforcement
'
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
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "brand_name_enforcement"

' -- Module-level brand rules dictionary ---------------------
' Key = correct form (String), Value = comma-separated incorrect variants (String)
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

' ============================================================
'  PRIVATE: Search for an incorrect variant and flag matches
' ============================================================
Private Sub SearchAndFlag(doc As Document, _
                           variant As String, _
                           correctForm As String, _
                           ByRef issues As Collection)
    Dim rng As Range
    Dim found As Boolean
    Dim finding As Object
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

        If EngineIsInPageRange(rng) Then
            On Error Resume Next
            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME, locStr, "Incorrect brand name:)
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
'  Format: one line per rule -- "CorrectForm=variant1,variant2"
' ============================================================
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

' ============================================================
'  PUBLIC: Load brand rules from a text file
'  Replaces existing rules with contents of the file.
'  Format: one line per rule -- "CorrectForm=variant1,variant2"
' ============================================================
Public Sub LoadBrandRules(filePath As String)
    Dim fileNum As Integer
    Dim lineText As String
    Dim eqPos As Long
    Dim correct As String
    Dim variants As String

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
    ' If file could not be loaded, fall back to defaults.
    ' Guard against Nothing before accessing .Count.
    If brandRules Is Nothing Then
        InitDefaultBrands
    ElseIf brandRules.Count = 0 Then
        InitDefaultBrands
    End If
End Sub

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for EngineIsInPageRange
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  PRIVATE: Late-bound wrapper for EngineGetLocationString
' ----------------------------------------------------------------

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
'  Late-bound wrapper: EngineIsInPageRange
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  Late-bound wrapper: EngineGetLocationString
' ----------------------------------------------------------------

' ----------------------------------------------------------------
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
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
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function
