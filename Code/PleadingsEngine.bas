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
'   - Rules_Spelling.bas        (spelling, licence/license, check/cheque)
'   - Rules_TextScan.bas        (repeated words, spell out under ten)
'   - Rules_Terms.bas           (custom term whitelist)
'   - Rules_NumberFormats.bas    (date/time, currency/number format)
'   - Rules_Punctuation.bas     (slash style, bracket integrity, dash usage)
'   - Rules_FootnoteIntegrity.bas (footnote integrity)
'   - Rules_Brands.bas          (brand name enforcement)
'   - Rules_FootnoteHarts.bas   (footnote Hart's rules)
'   - Rules_LegalTerms.bas      (mandated legal terms, always capitalise)
'   - Rules_Italics.bas         (anglicised terms, foreign names)
'   - Rules_Spacing.bas         (double spaces, commas, spacing)
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
Private nonEngTermPref As String   ' "ITALICS" or "REGULAR"
Private ruleErrorCount  As Long
Private ruleErrorLog    As String

' -- Precomputed page-filter boundaries (set once per run) --
Private gPageFilterEnabled   As Boolean
Private gPageFilterStartPos  As Long
Private gPageFilterEndPos    As Long

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
'  TARGET DOCUMENT SELECTION
'  Lets the user choose which open document to check.
'  Skips the macro host (ThisDocument) unless it's the only one.
' ============================================================
Public Function GetTargetDocument() As Document
    Set GetTargetDocument = Nothing
    Dim hostName As String
    On Error Resume Next
    hostName = ThisDocument.Name
    If Err.Number <> 0 Then hostName = "": Err.Clear
    On Error GoTo 0

    ' Build list of candidate documents (skip macro host)
    Dim candidates As New Collection
    Dim doc As Document
    For Each doc In Documents
        If doc.Name <> hostName Then
            candidates.Add doc
        End If
    Next doc

    ' No candidates: check if host is the only document
    If candidates.Count = 0 Then
        If Documents.Count = 1 Then
            ' Only the host is open -- confirm
            Dim hostChoice As VbMsgBoxResult
            hostChoice = MsgBox("The only open document is the macro host:" & vbCrLf & _
                                hostName & vbCrLf & vbCrLf & _
                                "Run the checker against this document?", _
                                vbYesNo + vbQuestion, "Pleadings Checker")
            If hostChoice = vbYes Then
                Set GetTargetDocument = Documents(1)
            End If
        Else
            MsgBox "No suitable target document is open." & vbCrLf & _
                   "Please open the document you want to check.", _
                   vbExclamation, "Pleadings Checker"
        End If
        Exit Function
    End If

    ' Single candidate: select automatically
    If candidates.Count = 1 Then
        Set GetTargetDocument = candidates(1)
        Debug.Print "GetTargetDocument: auto-selected " & candidates(1).Name
        Exit Function
    End If

    ' Multiple candidates: show picker
    Dim prompt As String
    prompt = "Select the document to check:" & vbCrLf & vbCrLf
    Dim idx As Long
    For idx = 1 To candidates.Count
        prompt = prompt & idx & ". " & candidates(idx).Name & vbCrLf
    Next idx

    Dim input As String
    input = InputBox(prompt, "Pleadings Checker - Select Document", "1")
    If Len(Trim(input)) = 0 Then Exit Function  ' cancelled

    Dim chosen As Long
    If IsNumeric(input) Then
        chosen = CLng(input)
        If chosen >= 1 And chosen <= candidates.Count Then
            Set GetTargetDocument = candidates(chosen)
        Else
            MsgBox "Invalid selection.", vbExclamation, "Pleadings Checker"
        End If
    Else
        MsgBox "Invalid selection.", vbExclamation, "Pleadings Checker"
    End If
End Function

' ============================================================
'  QUICK RUN (fallback when launcher is not imported)
'  Runs all available rules and shows summary via MsgBox.
' ============================================================
Public Sub RunQuick()
    TraceEnter "RunQuick"
    Dim targetDoc As Document
    Set targetDoc = GetTargetDocument()
    If targetDoc Is Nothing Then
        TraceExit "RunQuick", "no target selected"
        Exit Sub
    End If
    DebugLogDoc "RunQuick target", targetDoc
    Debug.Print "RunQuick: macro host=" & ThisDocument.Name & " target=" & targetDoc.Name

    Dim cfg As Object
    Set cfg = InitRuleConfig()
    SetPageRange 0, 0
    SetSpellingMode "UK"

    Dim issues As Collection
    Set issues = RunAllPleadingsRules(targetDoc, cfg)

    Dim summary As String
    summary = GetIssueSummary(issues)

    If issues.Count = 0 Then
        MsgBox "No issues found.", vbInformation, "Pleadings Checker"
    Else
        MsgBox summary, vbInformation, "Pleadings Checker"
        ApplySuggestionsAsTrackedChanges targetDoc, issues, True
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
'  NON-ENGLISH TERMS FORMAT PREFERENCE (italics or regular text)
' ============================================================
Public Sub SetNonEngTermPref(ByVal mode As String)
    nonEngTermPref = UCase(Trim(mode))
    If nonEngTermPref <> "REGULAR" Then nonEngTermPref = "ITALICS"
End Sub

Public Function GetNonEngTermPref() As String
    If Len(nonEngTermPref) = 0 Then nonEngTermPref = "ITALICS"
    GetNonEngTermPref = nonEngTermPref
End Function

' ============================================================
'  RULE CONFIGURATION
' ============================================================
Public Function InitRuleConfig() As Object
    Dim cfg As Object
    Set cfg = CreateObject("Scripting.Dictionary")

    cfg.Add "spellchecker", True
    cfg.Add "repeated_words", True
    cfg.Add "custom_term_whitelist", True
    cfg.Add "date_time_format", True
    cfg.Add "punctuation", True
    cfg.Add "currency_number_format", True
    cfg.Add "footnote_rules", True
    cfg.Add "brand_name_enforcement", True
    cfg.Add "mandated_legal_term_forms", True
    cfg.Add "always_capitalise_terms", True
    cfg.Add "non_english_terms", True
    cfg.Add "spell_out_under_ten", True
    cfg.Add "double_spaces", True

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
    Application.ScreenUpdating = False

    On Error GoTo RunnerCleanup

    ' -- Build paragraph position cache (one scan, enables O(log N) lookups) --
    BuildParagraphCache doc

    ' -- Precompute page-range character boundaries (one-time, cheap thereafter) --
    InitPageFilter doc

    ' -- Whitelist rule first (populates whitelistDict) --
    If IsRuleEnabled(config, "custom_term_whitelist") Then
        PerfTimerStart "custom_term_whitelist"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Terms.Check_CustomTermWhitelist", doc)
        PerfTimerEnd "custom_term_whitelist"
    End If
    DoEvents

    ' -- Spellchecker (spelling + licence/license + check/cheque) --
    If IsRuleEnabled(config, "spellchecker") Then
        PerfTimerStart "spellchecker"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_Spelling", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_LicenceLicense", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_CheckCheque", doc)
        PerfTimerEnd "spellchecker"
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
    ' -- Punctuation rules (single combined bucket) --
    If IsRuleEnabled(config, "punctuation") Then
        PerfTimerStart "punctuation"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Punctuation.Check_SlashStyle", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Punctuation.Check_BracketIntegrity", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Punctuation.Check_DashUsage", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Punctuation.Check_TriplicatePunctuation", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spacing.Check_DoubleCommas", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spacing.Check_SpaceBeforePunct", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spacing.Check_MissingSpaceAfterDot", doc)
        PerfTimerEnd "punctuation"
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
    ' -- Non-English term rules (italics) --
    If IsRuleEnabled(config, "non_english_terms") Then
        PerfTimerStart "non_english_terms"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Italics.Check_AnglicisedTermsNotItalic", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Italics.Check_ForeignNamesNotItalic", doc)
        PerfTimerEnd "non_english_terms"
    End If

RunnerCleanup:
    ' -- Restore application state (always runs) ----------------
    On Error Resume Next
    Application.ScreenUpdating = wasScreenUpdating
    Application.StatusBar = False   ' restore default status bar
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

    ' -- Hard safety guard: strip retired rule families -----------
    Set allIssues = FilterRetiredRules(allIssues)

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

    On Error GoTo HighlightCleanup

    Dim hlRS As Long, hlRE As Long

    For i = 1 To issues.Count
        Set finding = issues(i)

        ' Anchor validation gate
        If Not ValidateIssueAnchor(finding) Then
            TraceStep "ApplyHighlights", "INVALID ANCHOR SKIPPED i=" & i & _
                      " rule=" & CStr(GetIssueProp(finding, "RuleName"))
            GoTo NextHighlightIssue
        End If

        ' Read range into typed locals to avoid Variant coercion issues
        On Error Resume Next
        hlRS = CLng(GetIssueProp(finding, "RangeStart"))
        If Err.Number <> 0 Then hlRS = -1: Err.Clear
        hlRE = CLng(GetIssueProp(finding, "RangeEnd"))
        If Err.Number <> 0 Then hlRE = -1: Err.Clear
        On Error GoTo HighlightCleanup

        If hlRS >= 0 And hlRE > hlRS Then
            On Error Resume Next: Err.Clear
            Set rng = doc.Range(hlRS, hlRE)
            If Err.Number = 0 Then
                ' Apply yellow highlight to the flagged range
                rng.HighlightColorIndex = wdYellow
                If Err.Number <> 0 Then
                    DebugLogError "ApplyHighlights", "highlight i=" & i, Err.Number, Err.Description
                    Err.Clear
                End If
                If addComments Then
                    If ShouldCreateCommentForRule( _
                            CStr(GetIssueProp(finding, "RuleName")), finding) Then
                        TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                            "ApplyHighlights", "comment i=" & i
                    End If
                End If
            Else
                DebugLogError "ApplyHighlights", "doc.Range i=" & i & _
                    " start=" & hlRS & " end=" & hlRE, Err.Number, Err.Description
                Err.Clear
            End If
            On Error GoTo HighlightCleanup
        Else
            TraceStep "ApplyHighlights", "SKIPPED i=" & i & _
                      " -- invalid range start=" & hlRS & " end=" & hlRE
        End If
NextHighlightIssue:
    Next i

HighlightCleanup:
    On Error Resume Next
    Application.ScreenUpdating = wasScreenUpdating
    Application.StatusBar = False
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

    ' Enable tracking for the entire batch; restored once in cleanup.
    doc.TrackRevisions = True

    On Error GoTo TrackedCleanup

    ' Typed locals for each finding -- avoid repeated Variant coercion
    Dim rsVal As Long, reVal As Long
    Dim autoFix As Boolean
    Dim origStart As Long, origLen As Long
    Dim sugText As String
    Dim origText As String
    Dim skipAmendment As Boolean
    Dim chIdx As Long, ch As String
    Dim isOnlyWhitespace As Boolean
    Dim origHasPeriod As Boolean, sugHasPeriod As Boolean

    ' Get document story length for anchor validation
    Dim docStoryLen As Long
    On Error Resume Next
    docStoryLen = doc.Content.End
    If Err.Number <> 0 Then docStoryLen = -1: Err.Clear
    On Error GoTo TrackedCleanup

    ' Counters for debug summary
    Dim cntApplied As Long, cntCommentOnly As Long
    Dim cntSkippedAnchor As Long, cntSkippedUnsafe As Long
    cntApplied = 0: cntCommentOnly = 0
    cntSkippedAnchor = 0: cntSkippedUnsafe = 0

    ' Process from end of document backwards so tracked-change
    ' insertions / deletions do not shift positions of later issues
    For i = issues.Count To 1 Step -1
        Set finding = issues(i)

        ' Read range/autofix into typed locals
        On Error Resume Next
        rsVal = CLng(GetIssueProp(finding, "RangeStart"))
        If Err.Number <> 0 Then rsVal = -1: Err.Clear
        reVal = CLng(GetIssueProp(finding, "RangeEnd"))
        If Err.Number <> 0 Then reVal = -1: Err.Clear
        autoFix = False
        autoFix = CBool(GetIssueProp(finding, "AutoFixSafe"))
        If Err.Number <> 0 Then autoFix = False: Err.Clear
        On Error GoTo TrackedCleanup

        ' --- ANCHOR VALIDATION GATE ---
        If Not ValidateIssueAnchor(finding, docStoryLen) Then
            cntSkippedAnchor = cntSkippedAnchor + 1
            TraceStep "ApplyTrackedChanges", "INVALID ANCHOR SKIPPED i=" & i & _
                      " rule=" & CStr(GetIssueProp(finding, "RuleName")) & _
                      " start=" & rsVal & " end=" & reVal
            GoTo NextApplyIssue
        End If

        If rsVal >= 0 And reVal > rsVal Then
            On Error Resume Next: Err.Clear
            Set rng = doc.Range(rsVal, reVal)
            If Err.Number = 0 Then
                If autoFix Then
                    origStart = rng.Start
                    origLen = rng.End - rng.Start
                    ' Use ReplacementText only.  Suggestion is human-readable
                    ' prose and must NEVER be applied as literal replacement text.
                    ' An empty ReplacementText means "delete the range" -- distinct
                    ' from a MISSING key which means "no replacement available".
                    sugText = ""
                    If Not HasReplacementText(finding) Then
                        ' No machine-safe replacement -- skip amendment, add comment
                        TraceStep "ApplyTrackedChanges", "NO ReplacementText for i=" & i & _
                                  " rule=" & CStr(GetIssueProp(finding, "RuleName")) & "; comment-only"
                        If addComments And ShouldCreateCommentForRule( _
                                CStr(GetIssueProp(finding, "RuleName")), finding) Then
                            TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                                "ApplyTrackedChanges", "no-replacement-comment i=" & i
                        End If
                        GoTo NextApplyIssue
                    End If
                    sugText = CStr(GetIssueProp(finding, "ReplacementText"))

                    ' --- WHITESPACE VALIDATION GATE ---
                    origText = ""
                    origText = rng.Text
                    If Err.Number <> 0 Then origText = "": Err.Clear

                    skipAmendment = False

                    ' For deletions (empty suggestion = delete the range)
                    If Len(sugText) = 0 And Len(origText) > 0 Then
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
                                origHasPeriod = (InStr(1, origText, ".") > 0)
                                sugHasPeriod = (InStr(1, sugText, ".") > 0)
                                If origHasPeriod And Not sugHasPeriod Then
                                    skipAmendment = True
                                    Debug.Print "WHITESPACE VALIDATION: Skipped replacement '" & origText & "' -> '" & sugText & "' -- would remove period"
                                End If
                            End If
                        End If
                    End If

                    If skipAmendment Then
                        cntSkippedUnsafe = cntSkippedUnsafe + 1
                        TraceStep "ApplyTrackedChanges", "SKIPPED amendment i=" & i & _
                                  " orig=""" & Left$(origText, 30) & """ sug=""" & Left$(sugText, 30) & """"
                        If addComments And ShouldCreateCommentForRule( _
                                CStr(GetIssueProp(finding, "RuleName")), finding) Then
                            TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                                "ApplyTrackedChanges", "skip-comment i=" & i
                        End If
                        GoTo NextApplyIssue
                    End If

                    ' --- UNICODE SAFETY: reject replacement char U+FFFD ---
                    If Not IsReplacementSafe(sugText) Then
                        cntSkippedUnsafe = cntSkippedUnsafe + 1
                        TraceStep "ApplyTrackedChanges", "SKIPPED UNSAFE REPLACEMENT (U+FFFD) i=" & i
                        Debug.Print "UNICODE_SAFETY: replacement text contains U+FFFD for rule=" & _
                                    CStr(GetIssueProp(finding, "RuleName"))
                        If addComments And ShouldCreateCommentForRule( _
                                CStr(GetIssueProp(finding, "RuleName")), finding) Then
                            TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                                "ApplyTrackedChanges", "unsafe-replacement-comment i=" & i
                        End If
                        GoTo NextApplyIssue
                    End If

                    ' --- VERIFY EXACT MATCH before applying ---
                    ' If MatchedText was stored at detection time, verify the
                    ' document still contains it at this position.
                    Dim storedMatch As String
                    storedMatch = CStr(GetIssueProp(finding, "MatchedText"))
                    If Len(storedMatch) > 0 And Len(origText) > 0 Then
                        If origText <> storedMatch Then
                            cntSkippedUnsafe = cntSkippedUnsafe + 1
                            TraceStep "ApplyTrackedChanges", "SKIPPED STALE ANCHOR i=" & i & _
                                      " stored=""" & Left$(storedMatch, 30) & """ actual=""" & Left$(origText, 30) & """"
                            If addComments And ShouldCreateCommentForRule( _
                                    CStr(GetIssueProp(finding, "RuleName")), finding) Then
                                TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                                    "ApplyTrackedChanges", "stale-anchor-comment i=" & i
                            End If
                            GoTo NextApplyIssue
                        End If
                    End If

                    ' Apply tracked change
                    TraceStep "ApplyTrackedChanges", "APPLYING i=" & i & _
                              " range=" & origStart & "-" & (origStart + origLen) & _
                              " orig=""" & Left$(origText, 30) & """ -> """ & Left$(sugText, 30) & """"
                    TrySetRangeText rng, sugText, _
                        "ApplyTrackedChanges", "apply i=" & i
                    cntApplied = cntApplied + 1
                Else
                    cntCommentOnly = cntCommentOnly + 1
                    If addComments And ShouldCreateCommentForRule( _
                            CStr(GetIssueProp(finding, "RuleName")), finding) Then
                        TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                            "ApplyTrackedChanges", "comment-only i=" & i
                    End If
                End If
            Else
                DebugLogError "ApplyTrackedChanges", "doc.Range i=" & i & _
                    " start=" & rsVal & " end=" & reVal, Err.Number, Err.Description
                Err.Clear
            End If
NextApplyIssue:
            On Error GoTo TrackedCleanup
        Else
            TraceStep "ApplyTrackedChanges", "SKIPPED i=" & i & _
                      " -- invalid range start=" & rsVal & " end=" & reVal
        End If
    Next i

TrackedCleanup:
    ' Single cleanup path: always restore document and application state.
    On Error Resume Next
    doc.TrackRevisions = wasTrackingChanges
    Application.ScreenUpdating = wasScreenUpdating
    Application.StatusBar = False
    On Error GoTo 0
    TraceStep "ApplyTrackedChanges", "SUMMARY: applied=" & cntApplied & _
              " comment_only=" & cntCommentOnly & " skipped_anchor=" & cntSkippedAnchor & _
              " skipped_unsafe=" & cntSkippedUnsafe
    Debug.Print "ApplyTrackedChanges SUMMARY: applied=" & cntApplied & _
                " comment_only=" & cntCommentOnly & " skipped_anchor=" & cntSkippedAnchor & _
                " skipped_unsafe=" & cntSkippedUnsafe
    TraceExit "ApplyTrackedChanges"
End Sub

' ============================================================
'  PRIVATE: Build comment text from an issue dictionary
' ============================================================
Private Function BuildCommentText(ByVal finding As Object) As String
    Dim txt As String
    txt = GetIssueProp(finding, "Issue")

    ' Suppress "Suggestion:" tail for repeated-word findings
    Dim rn As String
    rn = LCase$(GetIssueProp(finding, "RuleName"))
    If rn = "repeated_words" Then
        BuildCommentText = txt
        Exit Function
    End If

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

    ' Resolve document name from explicit doc parameter
    Dim docName As String
    On Error Resume Next
    If Not doc Is Nothing Then
        docName = doc.Name
    Else
        docName = "(no document)"
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

    ' Count by confidence
    Dim confDict As Object
    Set confDict = CreateObject("Scripting.Dictionary")
    Dim invalidAnchorCount As Long
    invalidAnchorCount = 0
    For i = 1 To issues.Count
        Set finding = issues(i)
        Dim confLbl As String
        confLbl = CStr(GetIssueProp(finding, "ConfidenceLabel"))
        If Len(confLbl) = 0 Then confLbl = "unknown"
        If confDict.Exists(confLbl) Then
            confDict(confLbl) = confDict(confLbl) + 1
        Else
            confDict.Add confLbl, 1
        End If
        ' Check anchor validity
        If Not ValidateIssueAnchor(finding) Then
            invalidAnchorCount = invalidAnchorCount + 1
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
    Print #fileNum, "    },"
    Print #fileNum, "    ""counts_per_confidence"": {"
    Dim cKeys As Variant
    cKeys = confDict.keys
    For k = 0 To confDict.Count - 1
        If k < confDict.Count - 1 Then
            Print #fileNum, "      """ & EscJSON(CStr(cKeys(k))) & """: " & confDict(cKeys(k)) & ","
        Else
            Print #fileNum, "      """ & EscJSON(CStr(cKeys(k))) & """: " & confDict(cKeys(k))
        End If
    Next k
    Print #fileNum, "    },"
    Print #fileNum, "    ""invalid_anchor_count"": " & invalidAnchorCount
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
    If issues.Count = 0 Then
        GetIssueSummary = "No issues found."
        Exit Function
    End If

    ' Aggregate counts by UI label (not raw internal rule name)
    Dim uiCounts As Object
    Set uiCounts = CreateObject("Scripting.Dictionary")
    Dim finding As Object
    Dim i As Long
    Dim uiLbl As String

    For i = 1 To issues.Count
        Set finding = issues(i)
        uiLbl = GetUILabel(CStr(GetIssueProp(finding, "RuleName")))
        If uiCounts.Exists(uiLbl) Then
            uiCounts(uiLbl) = uiCounts(uiLbl) + 1
        Else
            uiCounts.Add uiLbl, 1
        End If
    Next i

    ' Sort by descending count using simple insertion sort
    Dim n As Long
    n = uiCounts.Count
    Dim sortLabels() As String
    Dim sortCounts() As Long
    ReDim sortLabels(0 To n - 1)
    ReDim sortCounts(0 To n - 1)
    Dim keys As Variant
    keys = uiCounts.keys
    Dim k As Long
    For k = 0 To n - 1
        sortLabels(k) = CStr(keys(k))
        sortCounts(k) = CLng(uiCounts(keys(k)))
    Next k

    Dim j As Long
    Dim tmpLbl As String
    Dim tmpCnt As Long
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If sortCounts(j) > sortCounts(i) Then
                tmpCnt = sortCounts(i): sortCounts(i) = sortCounts(j): sortCounts(j) = tmpCnt
                tmpLbl = sortLabels(i): sortLabels(i) = sortLabels(j): sortLabels(j) = tmpLbl
            End If
        Next j
    Next i

    ' Build output
    Dim result As String
    result = "Found " & issues.Count & " issue(s)." & vbCrLf & vbCrLf
    result = result & "By type:" & vbCrLf
    Dim maxShow As Long
    maxShow = 12
    If n < maxShow Then maxShow = n
    For k = 0 To maxShow - 1
        result = result & "  - " & sortLabels(k) & ": " & sortCounts(k) & vbCrLf
    Next k
    If n > maxShow Then
        result = result & "  + " & (n - maxShow) & " more" & vbCrLf
    End If

    ' Append slowest rules from profiler
    Dim slowest As String
    slowest = GetTopSlowestRules(3)
    If Len(slowest) > 0 Then
        result = result & vbCrLf & "Slowest: " & slowest
    End If

    GetIssueSummary = result
End Function

' ============================================================
'  RULE DISPLAY NAMES (for launcher summary)
' ============================================================
Public Function GetRuleDisplayNames() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")

    d.Add "spellchecker", "Spellchecker"
    d.Add "repeated_words", "Repeated Words"
    d.Add "custom_term_whitelist", "Custom Term Whitelist"
    d.Add "date_time_format", "Date/Time Format"
    d.Add "punctuation", "Punctuation Checker"
    d.Add "currency_number_format", "Currency/Number Formatting"
    d.Add "footnote_rules", "Footnote Rules"
    d.Add "brand_name_enforcement", "Custom Rules"
    d.Add "mandated_legal_term_forms", "Mandated Legal Terms"
    d.Add "always_capitalise_terms", "Always Capitalise Terms"
    d.Add "non_english_terms", "Non-English Terms"
    d.Add "spell_out_under_ten", "Spell Out Numbers Under 10"
    d.Add "double_spaces", "Double Spaces"

    Set GetRuleDisplayNames = d
End Function

' ============================================================
'  MAP INTERNAL RULE NAME TO USER-FACING LABEL
'  Sub-rules that belong to a combined bucket are mapped to
'  their parent UI label.  Unknown names are title-cased.
' ============================================================
Public Function GetUILabel(ByVal ruleName As String) As String
    Dim rn As String
    rn = LCase$(ruleName)

    ' Punctuation sub-rules -> "Punctuation"
    Select Case rn
        Case "slash_style", "bracket_integrity", "hyphens", "dash_usage", _
             "double_commas", "space_before_punct", "missing_space_after_dot", _
             "triplicate_punctuation", "punctuation"
            GetUILabel = "Punctuation Checker"
            Exit Function
        Case "spellchecker", "spelling", "licence_license", "check_cheque"
            GetUILabel = "Spellchecker"
            Exit Function
        Case "non_english_terms", "known_anglicised_terms_not_italic", _
             "foreign_names_not_italic"
            GetUILabel = "Non-English Terms"
            Exit Function
        Case "repeated_words"
            GetUILabel = "Repeated Words"
            Exit Function
        Case "double_spaces"
            GetUILabel = "Double Spaces"
            Exit Function
    End Select

    ' Fall back to display names dictionary
    Dim disp As Object
    Set disp = GetRuleDisplayNames()
    If disp.Exists(rn) Then
        GetUILabel = CStr(disp(rn))
    Else
        ' Title-case the rule name (replace underscores with spaces)
        Dim cleaned As String
        cleaned = Replace(rn, "_", " ")
        If Len(cleaned) > 0 Then
            Mid$(cleaned, 1, 1) = UCase$(Left$(cleaned, 1))
        End If
        GetUILabel = cleaned
    End If
End Function

' ============================================================
'  COMMENT SUPPRESSION: returns True if a comment bubble should
'  be created for this rule.  Returns False for trivial
'  whitespace/spacing rules that should be tracked-change only.
' ============================================================
Public Function ShouldCreateCommentForRule(ByVal ruleName As String, _
                                           Optional ByVal finding As Object = Nothing) As Boolean
    Dim rn As String
    rn = LCase$(ruleName)

    Select Case rn
        Case "double_spaces"
            ShouldCreateCommentForRule = False
            Exit Function
        Case "missing_space_after_dot"
            ShouldCreateCommentForRule = False
            Exit Function
        Case "trailing_spaces", "trailing_space"
            ShouldCreateCommentForRule = False
            Exit Function
    End Select

    ' Check issue text for spacing sub-types emitted under other rule names
    If Not finding Is Nothing Then
        On Error Resume Next
        Dim issText As String
        issText = ""
        If TypeName(finding) = "Dictionary" Then
            If finding.Exists("Issue") Then issText = LCase$(finding("Issue"))
        End If
        On Error GoTo 0
        If InStr(issText, "missing second space") > 0 Or _
           InStr(issText, "double space") > 0 Then
            ShouldCreateCommentForRule = False
            Exit Function
        End If
    End If

    ShouldCreateCommentForRule = True
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
' ============================================================
'  PRECOMPUTED PAGE FILTER
'  Called once per run after SetPageRangeFromString.
'  Translates page numbers into main-story character positions
'  so that rule inner loops need only cheap integer comparisons.
' ============================================================
Public Sub InitPageFilter(doc As Document)
    gPageFilterEnabled = False
    gPageFilterStartPos = 0
    gPageFilterEndPos = 0

    ' Nothing to filter?
    If pageRangeSet Is Nothing Then Exit Sub
    If pageRangeSet.Count = 0 Then Exit Sub

    ' Find min and max selected page numbers
    Dim minPage As Long, maxPage As Long
    Dim firstKey As Boolean
    firstKey = True
    Dim k As Variant
    For Each k In pageRangeSet.keys
        Dim pg As Long
        pg = CLng(k)
        If firstKey Then
            minPage = pg: maxPage = pg: firstKey = False
        Else
            If pg < minPage Then minPage = pg
            If pg > maxPage Then maxPage = pg
        End If
    Next k

    If minPage < 1 Then minPage = 1

    ' Determine total document pages
    Dim totalPages As Long
    On Error Resume Next
    totalPages = doc.ComputeStatistics(wdStatisticPages)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ' Cannot determine page count -- disable filter to be safe
        Exit Sub
    End If
    On Error GoTo 0

    ' Clamp maxPage to actual document length
    If maxPage > totalPages Then maxPage = totalPages
    If minPage > totalPages Then Exit Sub  ' selected pages beyond document

    ' Get start position: beginning of minPage
    Dim startRng As Range
    On Error Resume Next
    Set startRng = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=minPage)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    gPageFilterStartPos = startRng.Start

    ' Get end position: beginning of page after maxPage, or end of main story
    If maxPage >= totalPages Then
        ' Last page -- use end of main story
        gPageFilterEndPos = doc.Content.End
    Else
        Dim endRng As Range
        On Error Resume Next
        Set endRng = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=maxPage + 1)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            gPageFilterEndPos = doc.Content.End
        Else
            On Error GoTo 0
            gPageFilterEndPos = endRng.Start
        End If
    End If

    gPageFilterEnabled = True
    DebugLog "Page filter enabled: pages " & minPage & "-" & maxPage & _
             ", positions " & gPageFilterStartPos & "-" & gPageFilterEndPos
End Sub

' ============================================================
'  CHEAP OVERLAP HELPERS (integer comparisons only)
' ============================================================

' Returns True if a range overlaps the selected main-story span
' (or if page filtering is disabled).
Public Function IsInPageRange(rng As Range) As Boolean
    If Not gPageFilterEnabled Then
        IsInPageRange = True
        Exit Function
    End If
    ' Overlap test: range intersects [startPos, endPos)
    If rng.End <= gPageFilterStartPos Then
        IsInPageRange = False
    ElseIf rng.Start >= gPageFilterEndPos Then
        IsInPageRange = False
    Else
        IsInPageRange = True
    End If
End Function

' Position-based overlap check (no Range object needed)
Public Function IsInPageRangeByPos(ByVal startPos As Long, ByVal endPos As Long) As Boolean
    If Not gPageFilterEnabled Then
        IsInPageRangeByPos = True
        Exit Function
    End If
    If endPos <= gPageFilterStartPos Then
        IsInPageRangeByPos = False
    ElseIf startPos >= gPageFilterEndPos Then
        IsInPageRangeByPos = False
    Else
        IsInPageRangeByPos = True
    End If
End Function

' Returns True when startPos is past the end of the selected range.
' Rules iterating paragraphs in document order can use this to exit early.
Public Function IsPastPageFilter(ByVal startPos As Long) As Boolean
    If Not gPageFilterEnabled Then
        IsPastPageFilter = False
        Exit Function
    End If
    IsPastPageFilter = (startPos >= gPageFilterEndPos)
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
        ' Defensive filter: drop retired rule families
        If IsRetiredIssue(ruleIssues(i)) Then GoTo NextAddIssue
        master.Add ruleIssues(i)
NextAddIssue:
    Next i
End Sub

' ================================================================
'  PRIVATE: Returns True if an issue belongs to a retired rule
'  family (trailing spaces or after-heading spacing).  Used as a
'  belt-and-braces guard so legacy code can never emit these.
' ================================================================
Private Function IsRetiredIssue(ByVal item As Object) As Boolean
    IsRetiredIssue = False
    On Error Resume Next
    If TypeName(item) <> "Dictionary" Then Exit Function
    Dim rn As String
    rn = ""
    If item.Exists("RuleName") Then rn = LCase$(item("RuleName"))
    If Len(rn) = 0 Then Exit Function

    ' ---- Retired rule families (MVP pruning pass) ----

    ' Trailing spaces
    If rn = "trailing_spaces" Or rn = "trailing_space" Or _
       rn = "trailing whitespace" Then
        IsRetiredIssue = True: Exit Function
    End If

    ' Heading spacing (after/before heading)
    If rn = "after_heading_spacing" Or rn = "heading_spacing" Or _
       rn = "heading spacing" Or rn = "heading_spacing_consistency" Or _
       rn = "paragraph_break_consistency" Then
        IsRetiredIssue = True: Exit Function
    End If

    ' Font / formatting consistency
    If rn = "font_consistency" Or rn = "colour_formatting" Or _
       rn = "formatting_consistency" Then
        IsRetiredIssue = True: Exit Function
    End If

    ' Heading capitalisation / title formatting
    If rn = "heading_capitalisation" Or rn = "title_formatting" Then
        IsRetiredIssue = True: Exit Function
    End If

    ' Sequential numbering / clause number format
    If rn = "sequential_numbering" Or rn = "clause_number_format" Then
        IsRetiredIssue = True: Exit Function
    End If

    ' Defined terms
    If rn = "defined_terms" Or rn = "phrase_consistency" Then
        IsRetiredIssue = True: Exit Function
    End If

    ' Quote rules
    If rn = "quotation_mark_consistency" Or rn = "single_quotes_default" Or _
       rn = "smart_quote_consistency" Then
        IsRetiredIssue = True: Exit Function
    End If

    ' List rules
    If rn = "inline_list_format" Or rn = "list_punctuation" Or _
       rn = "list_rules" Then
        IsRetiredIssue = True: Exit Function
    End If

    ' ---- Message-based catch-all for edge cases ----
    Dim issText As String
    issText = ""
    If item.Exists("Issue") Then issText = LCase$(item("Issue"))

    If InStr(issText, "after-heading spacing") > 0 Or _
       InStr(issText, "after heading spacing") > 0 Or _
       InStr(issText, "spacing after heading") > 0 Or _
       InStr(issText, "font inconsistency") > 0 Or _
       InStr(issText, "dominant heading font") > 0 Or _
       InStr(issText, "dominant body font") > 0 Or _
       InStr(issText, "heading capitalisation") > 0 Or _
       InStr(issText, "title_case") > 0 Or _
       InStr(issText, "sentence_case") > 0 Or _
       InStr(issText, "numbering went backwards") > 0 Or _
       InStr(issText, "duplicate number") > 0 Or _
       InStr(issText, "defined term quote") > 0 Or _
       InStr(issText, "quotation mark consistency") > 0 Or _
       InStr(issText, "smart quote") > 0 Or _
       InStr(issText, "single quotes default") > 0 Or _
       InStr(issText, "double quotes default") > 0 Or _
       InStr(issText, "quote style consistency") > 0 Or _
       InStr(issText, "non-standard font colour") > 0 Then
        IsRetiredIssue = True: Exit Function
    End If
    On Error GoTo 0
End Function

' ================================================================
'  PRIVATE: Last-pass filter that strips any issues belonging to
'  retired rule families.  Called at the end of RunAllPleadingsRules
'  as a belt-and-braces guard for all downstream consumers.
' ================================================================
Private Function FilterRetiredRules(issues As Collection) As Collection
    Dim cleaned As New Collection
    Dim i As Long
    For i = 1 To issues.Count
        If Not IsRetiredIssue(issues(i)) Then
            cleaned.Add issues(i)
        End If
    Next i
    Set FilterRetiredRules = cleaned
End Function

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
                            Optional ByVal replacementText_ As String = "", _
                            Optional ByVal matchedText_ As String = "", _
                            Optional ByVal anchorKind_ As String = "exact_text", _
                            Optional ByVal confidenceLabel_ As String = "high", _
                            Optional ByVal sourceParagraphIndex_ As Long = 0) As Object
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
    If autoFixSafe_ Then d("ReplacementText") = replacementText_
    d("MatchedText") = matchedText_
    d("AnchorKind") = anchorKind_
    d("ConfidenceLabel") = confidenceLabel_
    d("SourceParagraphIndex") = sourceParagraphIndex_
    Set CreateIssue = d
End Function

' ================================================================
'  PUBLIC: Validate an issue anchor.  Returns True if the anchor
'  is plausible; False if the issue should be skipped or repaired.
' ================================================================
Public Function ValidateIssueAnchor(ByVal finding As Object, _
                                     Optional ByVal docStoryLen As Long = -1) As Boolean
    ValidateIssueAnchor = False
    On Error Resume Next
    Dim rs As Long, re As Long
    rs = CLng(GetIssueProp(finding, "RangeStart"))
    If Err.Number <> 0 Then rs = -1: Err.Clear
    re = CLng(GetIssueProp(finding, "RangeEnd"))
    If Err.Number <> 0 Then re = -1: Err.Clear
    On Error GoTo 0

    ' Basic validity
    If rs < 0 Then Exit Function
    If re <= rs Then Exit Function
    If docStoryLen > 0 And re > docStoryLen Then Exit Function

    ' Suspicious 1-char anchor for multi-word finding
    Dim issueLen As Long
    issueLen = re - rs
    If issueLen = 1 Then
        Dim issueText As String
        issueText = CStr(GetIssueProp(finding, "Issue"))
        Dim ak As String
        ak = CStr(GetIssueProp(finding, "AnchorKind"))
        ' Single-char anchor is OK for exact_text / token on a single char
        ' (e.g. a single quote or a single digit)
        ' but suspicious for paragraph-span issues
        If ak = "paragraph_span" Or ak = "paragraph_end" Then
            Debug.Print "ANCHOR_WARN: 1-char anchor for " & ak & " issue: " & Left$(issueText, 60)
        End If
    End If

    ValidateIssueAnchor = True
End Function

' ================================================================
'  PUBLIC: Check whether a replacement text contains the Unicode
'  replacement character U+FFFD.  If so, the replacement is unsafe.
' ================================================================
Public Function IsReplacementSafe(ByVal repText As String) As Boolean
    IsReplacementSafe = (InStr(1, repText, ChrW$(65533)) = 0)
End Function

' ================================================================
'  PRIVATE: Read a property from an finding (supports both
'  issue dictionary class and Dictionary-based issues)
' ================================================================
Private Function GetIssueProp(ByVal finding As Object, ByVal propName As String) As Variant
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
'  PRIVATE: Check whether a finding has a ReplacementText key
'  Distinguishes "key exists with empty value" (= delete) from
'  "key does not exist" (= no replacement available).
' ================================================================
Private Function HasReplacementText(ByVal finding As Object) As Boolean
    On Error Resume Next
    HasReplacementText = False
    If TypeName(finding) = "Dictionary" Then
        HasReplacementText = finding.Exists("ReplacementText")
    Else
        ' For non-dictionary objects, try to read the property
        Dim tmp As Variant
        tmp = CallByName(finding, "ReplacementText", VbGet)
        HasReplacementText = (Err.Number = 0)
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

' ================================================================
'  PRIVATE: Format an finding as JSON (supports both types)
' ================================================================
Private Function IssueToJSON(ByVal finding As Object) As String
    Dim s As String
    s = "    {" & vbCrLf
    s = s & "      ""rule"": """ & EscJSON(CStr(GetIssueProp(finding, "RuleName"))) & """," & vbCrLf
    s = s & "      ""location"": """ & EscJSON(CStr(GetIssueProp(finding, "Location"))) & """," & vbCrLf
    s = s & "      ""severity"": """ & EscJSON(CStr(GetIssueProp(finding, "Severity"))) & """," & vbCrLf
    s = s & "      ""finding"": """ & EscJSON(CStr(GetIssueProp(finding, "Issue"))) & """," & vbCrLf
    s = s & "      ""suggestion"": """ & EscJSON(CStr(GetIssueProp(finding, "Suggestion"))) & """," & vbCrLf
    ' Always emit replacement_text when the key exists (even if empty = deletion).
    ' HasReplacementText distinguishes "key present" from "key missing".
    If HasReplacementText(finding) Then
        Dim repText As String
        repText = CStr(GetIssueProp(finding, "ReplacementText"))
        s = s & "      ""replacement_text"": """ & EscJSON(repText) & """," & vbCrLf
    End If
    s = s & "      ""auto_fix_safe"": " & IIf(CBool(GetIssueProp(finding, "AutoFixSafe")), "true", "false") & "," & vbCrLf
    ' Enriched metadata
    Dim mt As String: mt = CStr(GetIssueProp(finding, "MatchedText"))
    If Len(mt) > 0 Then s = s & "      ""matched_text"": """ & EscJSON(Left$(mt, 80)) & """," & vbCrLf
    Dim ak As String: ak = CStr(GetIssueProp(finding, "AnchorKind"))
    If Len(ak) > 0 Then s = s & "      ""anchor_kind"": """ & EscJSON(ak) & """," & vbCrLf
    Dim cl As String: cl = CStr(GetIssueProp(finding, "ConfidenceLabel"))
    If Len(cl) > 0 Then s = s & "      ""confidence"": """ & EscJSON(cl) & """," & vbCrLf
    Dim spi As Long
    On Error Resume Next
    spi = CLng(GetIssueProp(finding, "SourceParagraphIndex"))
    If Err.Number <> 0 Then spi = 0: Err.Clear
    On Error GoTo 0
    If spi > 0 Then s = s & "      ""source_paragraph_index"": " & spi & "," & vbCrLf
    Dim rs As Long, re As Long
    On Error Resume Next
    rs = CLng(GetIssueProp(finding, "RangeStart"))
    If Err.Number <> 0 Then rs = 0: Err.Clear
    re = CLng(GetIssueProp(finding, "RangeEnd"))
    If Err.Number <> 0 Then re = 0: Err.Clear
    On Error GoTo 0
    s = s & "      ""range_start"": " & rs & "," & vbCrLf
    s = s & "      ""range_end"": " & re & vbCrLf
    s = s & "    }"
    IssueToJSON = s
End Function
