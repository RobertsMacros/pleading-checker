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
'   - Rules_Terms.bas           (Rules 5, 7, 23)
'   - Rules_Formatting.bas      (Rules 6, 11)
'   - Rules_NumberFormats.bas    (Rules 9, 18, 19)
'   - Rules_Lists.bas           (Rules 10, 15)
'   - Rules_Punctuation.bas     (Rules 14, 16)
'   - Rules_Quotes.bas          (Rules 17, 32, 33)
'   - Rules_FootnoteIntegrity.bas (Rule 20)
'   - Rules_Brands.bas          (Rule 22)
'   - Rules_FootnoteHarts.bas   (Rules 24, 25, 26, 27)
'   - Rules_LegalTerms.bas      (Rules 28, 29)
'   - Rules_Italics.bas         (Rules 30, 31)
'
' Installation:
'   1. Open the VBA Editor (Alt+F11)
'   2. Tools > References > check "Microsoft Scripting Runtime"
'   3. File > Import File > PleadingsEngine.bas
'   5. File > Import File > PleadingsLauncher.bas
'   6. Import whichever Rules_*.bas modules you need
'   7. Run the macro "PleadingsChecker"
' ============================================================
Option Explicit

' -- Module-level state --
Private ruleConfig      As Object
Private PAGE_RANGE_START As Long
Private PAGE_RANGE_END   As Long
Private whitelistDict   As Object
Private spellingMode    As String   ' "UK" or "US"

' ============================================================
'  ENTRY POINT
' ============================================================
Public Sub PleadingsChecker()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Pleadings Checker"
        Exit Sub
    End If
    ' Delegate to the launcher module (MsgBox/InputBox based, no form)
    On Error Resume Next
    Application.Run "PleadingsLauncher.LaunchChecker"
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ' Fallback: run all rules directly if launcher not imported
        RunQuick
    End If
    On Error GoTo 0
End Sub

' ============================================================
'  QUICK RUN (fallback when launcher is not imported)
'  Runs all available rules and shows summary via MsgBox.
' ============================================================
Public Sub RunQuick()
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
        ApplyHighlights ActiveDocument, issues, True
    End If
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
    cfg.Add "paragraph_break_consistency", True
    cfg.Add "defined_terms", True
    cfg.Add "clause_number_format", True
    cfg.Add "date_time_format", True
    cfg.Add "inline_list_format", True
    cfg.Add "font_consistency", True
    cfg.Add "licence_license", True
    cfg.Add "colour_formatting", True
    cfg.Add "slash_style", True
    cfg.Add "list_punctuation", True
    cfg.Add "bracket_integrity", True
    cfg.Add "quotation_mark_consistency", True
    cfg.Add "page_range", True
    cfg.Add "currency_number_format", True
    cfg.Add "footnote_integrity", True
    cfg.Add "title_formatting", True
    cfg.Add "brand_name_enforcement", True
    cfg.Add "phrase_consistency", True
    cfg.Add "footnotes_not_endnotes", True
    cfg.Add "footnote_terminal_full_stop", True
    cfg.Add "footnote_initial_capital", True
    cfg.Add "footnote_abbreviation_dictionary", True
    cfg.Add "mandated_legal_term_forms", True
    cfg.Add "always_capitalise_terms", True
    cfg.Add "known_anglicised_terms_not_italic", True
    cfg.Add "foreign_names_not_italic", True
    cfg.Add "single_quotes_default", True
    cfg.Add "smart_quote_consistency", True
    cfg.Add "spell_out_under_ten", True

    Set InitRuleConfig = cfg
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

    On Error Resume Next
    Set result = Application.Run(funcName, doc)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set TryRunRule = New Collection
        Exit Function
    End If
    On Error GoTo 0

    If result Is Nothing Then
        Set TryRunRule = New Collection
    Else
        Set TryRunRule = result
    End If
End Function

' ============================================================
'  MASTER RULE RUNNER
' ============================================================
Public Function RunAllPleadingsRules(doc As Document, _
                                     config As Object) As Collection
    Dim allIssues As New Collection
    Set ruleConfig = config

    ' -- Whitelist rule first (populates whitelistDict) --
    If IsRuleEnabled(config, "custom_term_whitelist") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Terms.Check_CustomTermWhitelist", doc)
    End If

    ' -- Spelling (bidirectional UK/US) --
    If IsRuleEnabled(config, "spelling") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_Spelling", doc)
    End If

    ' -- Text scanning rules --
    If IsRuleEnabled(config, "repeated_words") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_TextScan.Check_RepeatedWords", doc)
    End If

    If IsRuleEnabled(config, "spell_out_under_ten") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_TextScan.Check_SpellOutUnderTen", doc)
    End If

    ' -- Numbering rules --
    If IsRuleEnabled(config, "sequential_numbering") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Numbering.Check_SequentialNumbering", doc)
    End If

    If IsRuleEnabled(config, "clause_number_format") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Numbering.Check_ClauseNumberFormat", doc)
    End If

    ' -- Heading rules --
    If IsRuleEnabled(config, "heading_capitalisation") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Headings.Check_HeadingCapitalisation", doc)
    End If

    If IsRuleEnabled(config, "title_formatting") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Headings.Check_TitleFormatting", doc)
    End If

    ' -- Term rules --
    If IsRuleEnabled(config, "defined_terms") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Terms.Check_DefinedTerms", doc)
    End If

    If IsRuleEnabled(config, "phrase_consistency") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Terms.Check_PhraseConsistency", doc)
    End If

    ' -- Formatting rules --
    If IsRuleEnabled(config, "paragraph_break_consistency") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Formatting.Check_ParagraphBreakConsistency", doc)
    End If

    If IsRuleEnabled(config, "font_consistency") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Formatting.Check_FontConsistency", doc)
    End If

    ' -- Number format rules --
    If IsRuleEnabled(config, "date_time_format") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_NumberFormats.Check_DateTimeFormat", doc)
    End If

    If IsRuleEnabled(config, "page_range") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_NumberFormats.Check_PageRange", doc)
    End If

    If IsRuleEnabled(config, "currency_number_format") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_NumberFormats.Check_CurrencyNumberFormat", doc)
    End If

    ' -- List rules --
    If IsRuleEnabled(config, "inline_list_format") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Lists.Check_InlineListFormat", doc)
    End If

    If IsRuleEnabled(config, "list_punctuation") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Lists.Check_ListPunctuation", doc)
    End If

    ' -- UK/US variant rules (in Rules_Spelling) --
    If IsRuleEnabled(config, "licence_license") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_LicenceLicense", doc)
    End If

    If IsRuleEnabled(config, "colour_formatting") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_ColourFormatting", doc)
    End If

    ' -- Punctuation rules --
    If IsRuleEnabled(config, "slash_style") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Punctuation.Check_SlashStyle", doc)
    End If

    If IsRuleEnabled(config, "bracket_integrity") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Punctuation.Check_BracketIntegrity", doc)
    End If

    ' -- Quote rules --
    If IsRuleEnabled(config, "quotation_mark_consistency") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Quotes.Check_QuotationMarkConsistency", doc)
    End If

    If IsRuleEnabled(config, "single_quotes_default") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Quotes.Check_SingleQuotesDefault", doc)
    End If

    If IsRuleEnabled(config, "smart_quote_consistency") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Quotes.Check_SmartQuoteConsistency", doc)
    End If

    ' -- Footnote integrity --
    If IsRuleEnabled(config, "footnote_integrity") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_FootnoteIntegrity.Check_FootnoteIntegrity", doc)
    End If

    ' -- Brand names --
    If IsRuleEnabled(config, "brand_name_enforcement") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Brands.Check_BrandNameEnforcement", doc)
    End If

    ' -- Hart footnote rules --
    If IsRuleEnabled(config, "footnotes_not_endnotes") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_FootnoteHarts.Check_FootnotesNotEndnotes", doc)
    End If

    If IsRuleEnabled(config, "footnote_terminal_full_stop") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_FootnoteHarts.Check_FootnoteTerminalFullStop", doc)
    End If

    If IsRuleEnabled(config, "footnote_initial_capital") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_FootnoteHarts.Check_FootnoteInitialCapital", doc)
    End If

    If IsRuleEnabled(config, "footnote_abbreviation_dictionary") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_FootnoteHarts.Check_FootnoteAbbreviationDictionary", doc)
    End If

    ' -- Legal term rules --
    If IsRuleEnabled(config, "mandated_legal_term_forms") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_LegalTerms.Check_MandatedLegalTermForms", doc)
    End If

    If IsRuleEnabled(config, "always_capitalise_terms") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_LegalTerms.Check_AlwaysCapitaliseTerms", doc)
    End If

    ' -- Italic rules --
    If IsRuleEnabled(config, "known_anglicised_terms_not_italic") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Italics.Check_AnglicisedTermsNotItalic", doc)
    End If

    If IsRuleEnabled(config, "foreign_names_not_italic") Then
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Italics.Check_ForeignNamesNotItalic", doc)
    End If

    Set RunAllPleadingsRules = allIssues
End Function

' ============================================================
'  APPLY HIGHLIGHTS AND COMMENTS
' ============================================================
Public Sub ApplyHighlights(doc As Document, _
                           issues As Collection, _
                           Optional addComments As Boolean = True)
    Dim finding As Object
    Dim rng As Range
    Dim i As Long

    For i = 1 To issues.Count
        Set finding = issues(i)
        If GetIssueProp(finding, "RangeStart") >= 0 And GetIssueProp(finding, "RangeEnd") > GetIssueProp(finding, "RangeStart") Then
            On Error Resume Next: Err.Clear
            Set rng = doc.Range(GetIssueProp(finding, "RangeStart"), GetIssueProp(finding, "RangeEnd"))
            If Err.Number = 0 Then
                rng.HighlightColorIndex = wdYellow
                If addComments Then
                    doc.Comments.Add Range:=rng, _
                        Text:="[" & GetIssueProp(finding, "RuleName") & "] " & GetIssueProp(finding, "Issue") & _
                              " " & Chr(8212) & " Suggestion: " & GetIssueProp(finding, "Suggestion")
                End If
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

' ============================================================
'  APPLY SUGGESTIONS VIA TRACKED CHANGES
' ============================================================
Public Sub ApplySuggestionsAsTrackedChanges(doc As Document, _
                                             issues As Collection, _
                                             Optional addComments As Boolean = True)
    Dim finding As Object
    Dim rng As Range
    Dim i As Long
    Dim wasTrackingChanges As Boolean
    wasTrackingChanges = doc.TrackRevisions

    For i = 1 To issues.Count
        Set finding = issues(i)
        If GetIssueProp(finding, "RangeStart") >= 0 And GetIssueProp(finding, "RangeEnd") > GetIssueProp(finding, "RangeStart") Then
            On Error Resume Next: Err.Clear
            Set rng = doc.Range(GetIssueProp(finding, "RangeStart"), GetIssueProp(finding, "RangeEnd"))
            If Err.Number = 0 Then
                If GetIssueProp(finding, "AutoFixSafe") And Len(GetIssueProp(finding, "Suggestion")) > 0 Then
                    doc.TrackRevisions = True
                    rng.Text = GetIssueProp(finding, "Suggestion")
                    doc.TrackRevisions = wasTrackingChanges
                Else
                    rng.HighlightColorIndex = wdYellow
                    If addComments Then
                        doc.Comments.Add Range:=rng, _
                            Text:="[" & GetIssueProp(finding, "RuleName") & "] " & GetIssueProp(finding, "Issue") & _
                                  " " & Chr(8212) & " Suggestion: " & GetIssueProp(finding, "Suggestion")
                    End If
                End If
            End If
            On Error GoTo 0
        ElseIf addComments Then
            On Error Resume Next: Err.Clear
            If doc.Content.Start < doc.Content.End Then
                Set rng = doc.Range(doc.Content.Start, doc.Content.Start + 1)
                If Err.Number = 0 Then
                    doc.Comments.Add Range:=rng, _
                        Text:="[" & GetIssueProp(finding, "RuleName") & "] " & GetIssueProp(finding, "Issue") & _
                              " " & Chr(8212) & " Suggestion: " & GetIssueProp(finding, "Suggestion")
                End If
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

' ============================================================
'  GENERATE JSON REPORT
' ============================================================
Public Function GenerateReport(issues As Collection, _
                                filePath As String) As String
    Dim fileNum As Integer
    Dim finding As Object
    Dim i As Long

    fileNum = FreeFile
    Open filePath For Output As #fileNum

    Print #fileNum, "{"
    Print #fileNum, "  ""document"": """ & EscJSON(ActiveDocument.Name) & ""","
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
    d.Add "paragraph_break_consistency", "Paragraph Break Consistency"
    d.Add "defined_terms", "Defined Term Checker"
    d.Add "clause_number_format", "Clause Number Format"
    d.Add "date_time_format", "Date/Time Format Consistency"
    d.Add "inline_list_format", "Inline List Format"
    d.Add "font_consistency", "Font Consistency"
    d.Add "licence_license", "Licence/License Rule"
    d.Add "colour_formatting", "Colour Formatting Consistency"
    d.Add "slash_style", "Slash Style Checker"
    d.Add "list_punctuation", "List Punctuation Consistency"
    d.Add "bracket_integrity", "Bracket Integrity"
    d.Add "quotation_mark_consistency", "Quotation Mark Consistency"
    d.Add "page_range", "Page Range Filter"
    d.Add "currency_number_format", "Currency/Number Formatting"
    d.Add "footnote_integrity", "Footnote Integrity"
    d.Add "title_formatting", "Title Formatting Consistency"
    d.Add "brand_name_enforcement", "Brand Name Enforcement"
    d.Add "phrase_consistency", "Phrase Consistency"
    d.Add "footnotes_not_endnotes", "Footnotes Not Endnotes"
    d.Add "footnote_terminal_full_stop", "Footnote Terminal Full Stop"
    d.Add "footnote_initial_capital", "Footnote Initial Capital"
    d.Add "footnote_abbreviation_dictionary", "Footnote Abbreviation Dictionary"
    d.Add "mandated_legal_term_forms", "Mandated Legal Term Forms"
    d.Add "always_capitalise_terms", "Always Capitalise Terms"
    d.Add "known_anglicised_terms_not_italic", "Anglicised Terms Not Italic"
    d.Add "foreign_names_not_italic", "Foreign Names Not Italic"
    d.Add "single_quotes_default", "Single Quotes Default"
    d.Add "smart_quote_consistency", "Smart Quote Consistency"
    d.Add "spell_out_under_ten", "Spell Out Numbers Under 10"

    Set GetRuleDisplayNames = d
End Function

' ============================================================
'  HELPERS: PAGE RANGE
' ============================================================
Public Function IsInPageRange(rng As Range) As Boolean
    If PAGE_RANGE_START = 0 And PAGE_RANGE_END = 0 Then
        IsInPageRange = True
        Exit Function
    End If
    Dim pageNum As Long
    pageNum = rng.Information(wdActiveEndAdjustedPageNumber)
    IsInPageRange = (pageNum >= PAGE_RANGE_START And pageNum <= PAGE_RANGE_END)
End Function

Public Sub SetPageRange(startPage As Long, endPage As Long)
    PAGE_RANGE_START = startPage
    PAGE_RANGE_END = endPage
End Sub

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
    Dim para As Paragraph
    Dim paraIdx As Long

    On Error Resume Next
    pageNum = rng.Information(wdActiveEndAdjustedPageNumber)

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1
        If para.Range.Start >= rng.Start Then
            paraNum = paraIdx
            Exit For
        End If
    Next para
    If paraNum = 0 Then paraNum = paraIdx
    On Error GoTo 0

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
        ' Fall back to object property access
        Select Case propName
            Case "RuleName":    GetIssueProp = GetIssueProp(finding, "RuleName")
            Case "Location":    GetIssueProp = GetIssueProp(finding, "Location")
            Case "Issue":       GetIssueProp = GetIssueProp(finding, "Issue")
            Case "Suggestion":  GetIssueProp = GetIssueProp(finding, "Suggestion")
            Case "Severity":    GetIssueProp = GetIssueProp(finding, "Severity")
            Case "RangeStart":  GetIssueProp = GetIssueProp(finding, "RangeStart")
            Case "RangeEnd":    GetIssueProp = GetIssueProp(finding, "RangeEnd")
            Case "AutoFixSafe": GetIssueProp = GetIssueProp(finding, "AutoFixSafe")
        End Select
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
    s = s & "      ""auto_fix_safe"": " & IIf(CBool(GetIssueProp(finding, "AutoFixSafe")), "true", "false") & vbCrLf
    s = s & "    }"
    IssueToJSON = s
End Function
