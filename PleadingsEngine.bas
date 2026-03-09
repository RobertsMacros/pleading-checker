Attribute VB_Name = "PleadingsEngine"
' ============================================================
' PleadingsEngine.bas
' Core engine for the Pleadings Checker rule system.
' Coordinates running 34 proofreading rules against a Word
' document, collecting structured PleadingsIssue results,
' applying highlights/comments/tracked-changes, and generating
' JSON reports.
'
' Dependencies:
'   - PleadingsIssue.cls   (structured result class)
'   - PleadingsRules.bas   (individual rule implementations)
'   - frmPleadingsChecker   (user form)
'   - Microsoft Scripting Runtime (Tools > References)
'
' Installation:
'   1. Open the VBA Editor (Alt+F11)
'   2. Tools > References > check "Microsoft Scripting Runtime"
'   3. File > Import File > select PleadingsEngine.bas
'   4. File > Import File > select PleadingsIssue.cls
'   5. File > Import File > select PleadingsRules.bas
'   6. File > Import File > select frmPleadingsChecker.frm
'   7. Run the macro "PleadingsChecker" (or assign to a ribbon button)
' ============================================================
Option Explicit

' ── Module-level state ────────────────────────────────────────
Private ruleConfig      As Scripting.Dictionary   ' rule name (String) -> enabled (Boolean)
Private PAGE_RANGE_START As Long                  ' 0 = no restriction
Private PAGE_RANGE_END   As Long                  ' 0 = no restriction
Private whitelistDict   As Scripting.Dictionary   ' custom term whitelist

' ════════════════════════════════════════════════════════════
'  ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Sub PleadingsChecker()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Pleadings Checker"
        Exit Sub
    End If
    frmPleadingsChecker.Show
End Sub

' ════════════════════════════════════════════════════════════
'  RULE CONFIGURATION
'  Creates a Dictionary with all 34 rule names defaulting
'  to True (enabled). The form can toggle individual rules.
' ════════════════════════════════════════════════════════════
Public Function InitRuleConfig() As Scripting.Dictionary
    Dim cfg As New Scripting.Dictionary

    cfg.Add "british_spelling", True
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

    ' ── Bucket 1: Hart rules ──────────────────────────────────
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

' ════════════════════════════════════════════════════════════
'  MASTER RULE RUNNER
'  Iterates every enabled rule, calls the corresponding
'  function from PleadingsRules, and collects all
'  PleadingsIssue objects into a single Collection.
'  Each rule call is wrapped in error handling so that
'  one failure never stops the remaining rules.
' ════════════════════════════════════════════════════════════
Public Function RunAllPleadingsRules(doc As Document, _
                                     config As Scripting.Dictionary) As Collection
    Dim allIssues As New Collection
    Dim ruleIssues As Collection
    Dim issue As PleadingsIssue
    Dim i As Long

    ' Store config at module level for helper access
    Set ruleConfig = config

    ' ── Run whitelist rule first (populates whitelistDict) ──
    If config.Exists("custom_term_whitelist") Then
        If config("custom_term_whitelist") = True Then
            On Error Resume Next: Err.Clear
            Set ruleIssues = Check_CustomTermWhitelist(doc)
            If Err.Number = 0 Then
                If Not ruleIssues Is Nothing Then
                    For i = 1 To ruleIssues.Count
                        allIssues.Add ruleIssues(i)
                    Next i
                End If
            End If
            On Error GoTo 0
        End If
    End If

    ' ── Run remaining rules ─────────────────────────────────

    ' british_spelling
    If IsRuleEnabled(config, "british_spelling") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_BritishSpelling(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' repeated_words
    If IsRuleEnabled(config, "repeated_words") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_RepeatedWords(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' sequential_numbering
    If IsRuleEnabled(config, "sequential_numbering") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_SequentialNumbering(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' heading_capitalisation
    If IsRuleEnabled(config, "heading_capitalisation") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_HeadingCapitalisation(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' paragraph_break_consistency
    If IsRuleEnabled(config, "paragraph_break_consistency") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_ParagraphBreakConsistency(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' defined_terms
    If IsRuleEnabled(config, "defined_terms") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_DefinedTerms(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' clause_number_format
    If IsRuleEnabled(config, "clause_number_format") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_ClauseNumberFormat(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' date_time_format
    If IsRuleEnabled(config, "date_time_format") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_DateTimeFormat(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' inline_list_format
    If IsRuleEnabled(config, "inline_list_format") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_InlineListFormat(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' font_consistency
    If IsRuleEnabled(config, "font_consistency") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_FontConsistency(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' licence_license
    If IsRuleEnabled(config, "licence_license") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_LicenceLicense(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' colour_formatting
    If IsRuleEnabled(config, "colour_formatting") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_ColourFormatting(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' slash_style
    If IsRuleEnabled(config, "slash_style") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_SlashStyle(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' list_punctuation
    If IsRuleEnabled(config, "list_punctuation") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_ListPunctuation(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' bracket_integrity
    If IsRuleEnabled(config, "bracket_integrity") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_BracketIntegrity(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' quotation_mark_consistency
    If IsRuleEnabled(config, "quotation_mark_consistency") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_QuotationMarkConsistency(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' page_range
    If IsRuleEnabled(config, "page_range") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_PageRange(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' currency_number_format
    If IsRuleEnabled(config, "currency_number_format") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_CurrencyNumberFormat(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' footnote_integrity
    If IsRuleEnabled(config, "footnote_integrity") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_FootnoteIntegrity(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' title_formatting
    If IsRuleEnabled(config, "title_formatting") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_TitleFormatting(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' brand_name_enforcement
    If IsRuleEnabled(config, "brand_name_enforcement") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_BrandNameEnforcement(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' phrase_consistency
    If IsRuleEnabled(config, "phrase_consistency") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_PhraseConsistency(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' ── Bucket 1: Hart rules ──────────────────────────────────

    ' footnotes_not_endnotes
    If IsRuleEnabled(config, "footnotes_not_endnotes") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_FootnotesNotEndnotes(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' footnote_terminal_full_stop
    If IsRuleEnabled(config, "footnote_terminal_full_stop") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_FootnoteTerminalFullStop(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' footnote_initial_capital
    If IsRuleEnabled(config, "footnote_initial_capital") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_FootnoteInitialCapital(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' footnote_abbreviation_dictionary
    If IsRuleEnabled(config, "footnote_abbreviation_dictionary") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_FootnoteAbbreviationDictionary(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' mandated_legal_term_forms
    If IsRuleEnabled(config, "mandated_legal_term_forms") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_MandatedLegalTermForms(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' always_capitalise_terms
    If IsRuleEnabled(config, "always_capitalise_terms") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_AlwaysCapitaliseTerms(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' known_anglicised_terms_not_italic
    If IsRuleEnabled(config, "known_anglicised_terms_not_italic") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_AnglicisedTermsNotItalic(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' foreign_names_not_italic
    If IsRuleEnabled(config, "foreign_names_not_italic") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_ForeignNamesNotItalic(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' single_quotes_default
    If IsRuleEnabled(config, "single_quotes_default") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_SingleQuotesDefault(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' smart_quote_consistency
    If IsRuleEnabled(config, "smart_quote_consistency") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_SmartQuoteConsistency(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    ' spell_out_under_ten
    If IsRuleEnabled(config, "spell_out_under_ten") Then
        On Error Resume Next: Err.Clear
        Set ruleIssues = Check_SpellOutUnderTen(doc)
        If Err.Number = 0 Then AddIssuesToCollection allIssues, ruleIssues
        On Error GoTo 0
    End If

    Set RunAllPleadingsRules = allIssues
End Function

' ════════════════════════════════════════════════════════════
'  APPLY HIGHLIGHTS AND COMMENTS
'  Loops all issues and marks them in the document with
'  yellow highlighting and optional review comments.
' ════════════════════════════════════════════════════════════
Public Sub ApplyHighlights(doc As Document, _
                           issues As Collection, _
                           Optional addComments As Boolean = True)
    Dim issue As PleadingsIssue
    Dim rng As Range
    Dim i As Long

    For i = 1 To issues.Count
        Set issue = issues(i)

        ' Skip issues without valid range positions
        If issue.RangeStart >= 0 And issue.RangeEnd > issue.RangeStart Then
            On Error Resume Next: Err.Clear
            Set rng = doc.Range(issue.RangeStart, issue.RangeEnd)
            If Err.Number = 0 Then
                ' Apply yellow highlight
                rng.HighlightColorIndex = wdYellow

                ' Add review comment if requested
                If addComments Then
                    doc.Comments.Add Range:=rng, _
                        Text:="[" & issue.RuleName & "] " & issue.Issue & _
                              " " & Chr(8212) & " Suggestion: " & issue.Suggestion
                End If
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  GENERATE JSON REPORT
'  Writes a structured JSON file with all issues, summary
'  counts, and most-frequent-types ranking.
'  Returns a summary string for display in the form.
' ════════════════════════════════════════════════════════════
Public Function GenerateReport(issues As Collection, _
                                filePath As String) As String
    Dim fileNum As Integer
    Dim issue As PleadingsIssue
    Dim i As Long
    Dim summaryStr As String

    fileNum = FreeFile
    Open filePath For Output As #fileNum

    ' ── Document header ─────────────────────────────────────
    Print #fileNum, "{"
    Print #fileNum, "  ""document"": """ & EscJSON(ActiveDocument.Name) & ""","
    Print #fileNum, "  ""timestamp"": """ & Format(Now, "yyyy-mm-ddThh:nn:ss") & ""","
    Print #fileNum, "  ""total_issues"": " & issues.Count & ","

    ' ── Issues array ────────────────────────────────────────
    Print #fileNum, "  ""issues"": ["
    For i = 1 To issues.Count
        Set issue = issues(i)
        If i < issues.Count Then
            Print #fileNum, issue.ToJSON() & ","
        Else
            Print #fileNum, issue.ToJSON()
        End If
    Next i
    Print #fileNum, "  ],"

    ' ── Summary: counts per rule ────────────────────────────
    Dim countDict As New Scripting.Dictionary
    For i = 1 To issues.Count
        Set issue = issues(i)
        If countDict.Exists(issue.RuleName) Then
            countDict(issue.RuleName) = countDict(issue.RuleName) + 1
        Else
            countDict.Add issue.RuleName, 1
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

    ' ── Most frequent types (sorted descending by count) ───
    Print #fileNum, "    ""most_frequent_types"": ["

    ' Simple selection sort on keys by count descending
    Dim sortedKeys() As String
    Dim sortedCounts() As Long
    Dim nRules As Long
    nRules = countDict.Count

    If nRules > 0 Then
        ReDim sortedKeys(0 To nRules - 1)
        ReDim sortedCounts(0 To nRules - 1)

        For k = 0 To nRules - 1
            sortedKeys(k) = CStr(keys(k))
            sortedCounts(k) = countDict(keys(k))
        Next k

        ' Selection sort descending
        Dim j As Long
        Dim maxIdx As Long
        Dim tmpStr As String
        Dim tmpLng As Long
        For k = 0 To nRules - 2
            maxIdx = k
            For j = k + 1 To nRules - 1
                If sortedCounts(j) > sortedCounts(maxIdx) Then
                    maxIdx = j
                End If
            Next j
            If maxIdx <> k Then
                tmpStr = sortedKeys(k)
                sortedKeys(k) = sortedKeys(maxIdx)
                sortedKeys(maxIdx) = tmpStr
                tmpLng = sortedCounts(k)
                sortedCounts(k) = sortedCounts(maxIdx)
                sortedCounts(maxIdx) = tmpLng
            End If
        Next k

        ' Write sorted entries
        For k = 0 To nRules - 1
            Dim comma As String
            If k < nRules - 1 Then comma = "," Else comma = ""
            Print #fileNum, "      { ""rule"": """ & EscJSON(sortedKeys(k)) & _
                            """, ""count"": " & sortedCounts(k) & " }" & comma
        Next k
    End If

    Print #fileNum, "    ]"
    Print #fileNum, "  }"
    Print #fileNum, "}"

    Close #fileNum

    ' ── Build summary string for display ────────────────────
    summaryStr = "Report saved: " & filePath & vbCrLf
    summaryStr = summaryStr & "Total issues: " & issues.Count & vbCrLf
    If nRules > 0 Then
        summaryStr = summaryStr & vbCrLf & "Most frequent:" & vbCrLf
        For k = 0 To nRules - 1
            summaryStr = summaryStr & "  " & sortedKeys(k) & ": " & sortedCounts(k) & vbCrLf
        Next k
    End If

    GenerateReport = summaryStr
End Function

' ════════════════════════════════════════════════════════════
'  HUMAN-READABLE ISSUE SUMMARY
'  Builds a multi-line string with counts per rule.
' ════════════════════════════════════════════════════════════
Public Function GetIssueSummary(issues As Collection) As String
    Dim countDict As New Scripting.Dictionary
    Dim issue As PleadingsIssue
    Dim i As Long

    ' Count issues per rule
    For i = 1 To issues.Count
        Set issue = issues(i)
        If countDict.Exists(issue.RuleName) Then
            countDict(issue.RuleName) = countDict(issue.RuleName) + 1
        Else
            countDict.Add issue.RuleName, 1
        End If
    Next i

    ' Format output lines
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
        result = result & CStr(keys(k)) & ": " & cnt & " issue"
        If cnt <> 1 Then result = result & "s"
        result = result & vbCrLf
    Next k

    result = result & vbCrLf & "Total: " & issues.Count & " issue"
    If issues.Count <> 1 Then result = result & "s"

    GetIssueSummary = result
End Function

' ════════════════════════════════════════════════════════════
'  HELPER: PAGE RANGE FILTER
'  Returns True if the range falls within the configured
'  page restriction, or if no restriction is set (both 0).
' ════════════════════════════════════════════════════════════
Public Function IsInPageRange(rng As Range) As Boolean
    ' No restriction when both bounds are zero
    If PAGE_RANGE_START = 0 And PAGE_RANGE_END = 0 Then
        IsInPageRange = True
        Exit Function
    End If

    Dim pageNum As Long
    pageNum = rng.Information(wdActiveEndAdjustedPageNumber)

    IsInPageRange = (pageNum >= PAGE_RANGE_START And pageNum <= PAGE_RANGE_END)
End Function

' ════════════════════════════════════════════════════════════
'  HELPER: WHITELIST LOOKUP
'  Case-insensitive check against the custom term whitelist.
' ════════════════════════════════════════════════════════════
Public Function IsWhitelistedTerm(term As String) As Boolean
    If whitelistDict Is Nothing Then
        IsWhitelistedTerm = False
        Exit Function
    End If

    IsWhitelistedTerm = whitelistDict.Exists(LCase(term))
End Function

' ════════════════════════════════════════════════════════════
'  HELPER: SET PAGE RANGE
' ════════════════════════════════════════════════════════════
Public Sub SetPageRange(startPage As Long, endPage As Long)
    PAGE_RANGE_START = startPage
    PAGE_RANGE_END = endPage
End Sub

' ════════════════════════════════════════════════════════════
'  HELPER: SET WHITELIST
' ════════════════════════════════════════════════════════════
Public Sub SetWhitelist(terms As Scripting.Dictionary)
    Set whitelistDict = terms
End Sub

' ════════════════════════════════════════════════════════════
'  HELPER: LOCATION STRING
'  Returns "page N paragraph M" for a given range.
' ════════════════════════════════════════════════════════════
Public Function GetLocationString(rng As Range, doc As Document) As String
    Dim pageNum As Long
    Dim paraNum As Long

    On Error Resume Next

    ' Page number from Word's adjusted page counter
    pageNum = rng.Information(wdActiveEndAdjustedPageNumber)

    ' Paragraph number via single COM call — Word counts internally,
    ' vastly faster than iterating For Each para in VBA (which was
    ' O(n) per call and caused minute-long runtimes on large docs)
    paraNum = doc.Range(doc.Content.Start, rng.Start).Paragraphs.Count
    If Err.Number <> 0 Then paraNum = 0: Err.Clear

    On Error GoTo 0

    GetLocationString = "page " & pageNum & " paragraph " & paraNum
End Function

' ════════════════════════════════════════════════════════════
'  APPLY SUGGESTIONS VIA TRACKED CHANGES
'  For issues flagged as auto-fix safe, applies the suggestion
'  text using Word's tracked changes so the user can accept
'  or reject each change individually.
'  For non-auto-fix issues, adds a comment with the suggestion.
' ════════════════════════════════════════════════════════════
Public Sub ApplySuggestionsAsTrackedChanges(doc As Document, _
                                             issues As Collection, _
                                             Optional addComments As Boolean = True)
    Dim issue As PleadingsIssue
    Dim rng As Range
    Dim i As Long
    Dim wasTrackingChanges As Boolean

    ' Remember current tracking state
    wasTrackingChanges = doc.TrackRevisions

    For i = 1 To issues.Count
        Set issue = issues(i)

        ' Skip issues without valid range positions
        If issue.RangeStart >= 0 And issue.RangeEnd > issue.RangeStart Then
            On Error Resume Next: Err.Clear
            Set rng = doc.Range(issue.RangeStart, issue.RangeEnd)
            If Err.Number = 0 Then
                If issue.AutoFixSafe And Len(issue.Suggestion) > 0 Then
                    ' Apply replacement via tracked changes
                    doc.TrackRevisions = True
                    rng.Text = issue.Suggestion
                    doc.TrackRevisions = wasTrackingChanges
                Else
                    ' Highlight and comment for suggest-only issues
                    rng.HighlightColorIndex = wdYellow
                    If addComments Then
                        doc.Comments.Add Range:=rng, _
                            Text:="[" & issue.RuleName & "] " & issue.Issue & _
                                  " " & Chr(8212) & " Suggestion: " & issue.Suggestion
                    End If
                End If
            End If
            On Error GoTo 0
        ElseIf addComments Then
            ' Document-level issues with no range: add comment at start
            On Error Resume Next: Err.Clear
            If doc.Content.Start < doc.Content.End Then
                Set rng = doc.Range(doc.Content.Start, doc.Content.Start + 1)
                If Err.Number = 0 Then
                    doc.Comments.Add Range:=rng, _
                        Text:="[" & issue.RuleName & "] " & issue.Issue & _
                              " " & Chr(8212) & " Suggestion: " & issue.Suggestion
                End If
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  RULE METADATA
'  Returns a Dictionary mapping rule_name -> display label
'  for use by the form's dynamic rule list.
' ════════════════════════════════════════════════════════════
Public Function GetRuleDisplayNames() As Scripting.Dictionary
    Dim d As New Scripting.Dictionary

    d.Add "british_spelling", "British Spelling Enforcement"
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

' ════════════════════════════════════════════════════════════
'  PRIVATE HELPERS
' ════════════════════════════════════════════════════════════

' Checks if a rule is enabled in the config dictionary
Private Function IsRuleEnabled(config As Scripting.Dictionary, _
                                ruleName As String) As Boolean
    If config.Exists(ruleName) Then
        IsRuleEnabled = CBool(config(ruleName))
    Else
        IsRuleEnabled = False
    End If
End Function

' Transfers issues from a rule-level collection into the master collection
Private Sub AddIssuesToCollection(master As Collection, _
                                   ruleIssues As Collection)
    Dim i As Long
    If ruleIssues Is Nothing Then Exit Sub
    For i = 1 To ruleIssues.Count
        master.Add ruleIssues(i)
    Next i
End Sub

' Escapes special characters for safe JSON string output
Private Function EscJSON(ByVal txt As String) As String
    txt = Replace(txt, "\", "\\")
    txt = Replace(txt, """", "\""")
    txt = Replace(txt, vbCr, "\r")
    txt = Replace(txt, vbLf, "\n")
    txt = Replace(txt, vbTab, "\t")
    EscJSON = txt
End Function
