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

' -- Cooperative cancellation --
Public gCancelRun       As Boolean
Private Const ERR_RUN_CANCELLED As Long = vbObjectError + 513

' -- Finding output mode constants --
Public Const OUTPUT_TRACKED_SAFE As String = "TRACKED_SAFE"
Public Const OUTPUT_COMMENT_ONLY As String = "COMMENT_ONLY"
Public Const OUTPUT_REPORT_ONLY As String = "REPORT_ONLY"
Public Const OUTPUT_GROUPED_REPORT As String = "GROUPED_REPORT"

' -- Comment spam / grouped report thresholds --
Private Const SPELLING_COMMENT_THRESHOLD As Long = 25
Private Const FOOTNOTE_COMMENT_THRESHOLD As Long = 15
Private Const MAX_INLINE_SPELLING_COMMENTS As Long = 20
Private Const MAX_INLINE_FOOTNOTE_COMMENTS As Long = 10
Private Const MAX_INLINE_TOTAL_COMMENTS As Long = 80
Private Const MAX_DUPLICATE_COMMENT_TEXT_PER_RUN As Long = 3
Private gSpellingGroups     As Object   ' Dictionary: "wrong|right" -> count
Private gSpellingExamples   As Object   ' Dictionary: "wrong|right" -> first 3 locations
Private gFootnoteGroups     As Object   ' Dictionary: "issueKey" -> Dictionary with Count, NoteNumbers, IssueText
Private gGroupedSpellingCount As Long
Private gGroupedFootnoteCount As Long
Private gCommentsCreated    As Long
Private gTrackedEditsApplied As Long
Private gEditsSkippedUnsafe As Long

' -- Unsafe autofix category set (report-only, no tracked edits) --
Private gUnsafeAutofixRules As Object   ' Dictionary of rule names -> True

' -- Tracked-safe allow-list (narrow: only rules proven safe for auto-fix) --
Private gTrackedSafeRules As Object     ' Dictionary of rule names -> True

' -- Comment-safe allow-list (rules that may create inline comments) --
Private gCommentSafeRules As Object     ' Dictionary of rule names -> True

' -- Document complexity state (computed once per run) --
Private gDocIsComplex As Boolean

' -- Duplicate comment text tracking (suppress repeated identical comments) --
Private gCommentTextCounts As Object    ' Dictionary: comment_text -> count

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
Private pageRangeString As String  ' Raw user-entered page-range spec (preserved for reports)
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

' -- Block-quote / TOC / cover-page region cache (built once per run) --
Private bqStarts()      As Long
Private bqEnds()        As Long
Private bqCount         As Long
Private tocStarts()     As Long
Private tocEnds()       As Long
Private tocCount        As Long
Private coverPageEnd    As Long   ' -1 = no cover page detected

' ============================================================
'  COOPERATIVE CANCELLATION HELPERS
' ============================================================
Public Sub ResetCancelRun()
    gCancelRun = False
End Sub

Public Sub RequestCancelRun()
    gCancelRun = True
End Sub

Public Function CancelRunRequested() As Boolean
    CancelRunRequested = gCancelRun
End Function

Private Sub CheckCancellation()
    DoEvents
    If gCancelRun Then Err.Raise ERR_RUN_CANCELLED, "PleadingsEngine", "Run cancelled"
End Sub

' Public macro target for cooperative cancellation.
' Must be a parameterless Public Sub so Word can call it by name.
Public Sub CancelCurrentRun()
    RequestCancelRun
End Sub

' ============================================================
'  GROUPED REPORT + COMMENT SUPPRESSION HELPERS
' ============================================================
Private Sub InitGroupedReportState()
    Set gSpellingGroups = CreateObject("Scripting.Dictionary")
    Set gSpellingExamples = CreateObject("Scripting.Dictionary")
    Set gFootnoteGroups = CreateObject("Scripting.Dictionary")
    gGroupedSpellingCount = 0
    gGroupedFootnoteCount = 0
    gCommentsCreated = 0
    gTrackedEditsApplied = 0
    gEditsSkippedUnsafe = 0
    gDocIsComplex = False

    ' Duplicate comment text tracking
    Set gCommentTextCounts = CreateObject("Scripting.Dictionary")

    ' -- Unsafe autofix rules (report-only, never tracked-edit) --
    Set gUnsafeAutofixRules = CreateObject("Scripting.Dictionary")
    gUnsafeAutofixRules("missing_space_after_dot") = True
    gUnsafeAutofixRules("space_before_punct") = True
    gUnsafeAutofixRules("hyphens") = True
    gUnsafeAutofixRules("dash_usage") = True
    gUnsafeAutofixRules("footnote_integrity") = True
    gUnsafeAutofixRules("duplicate_footnotes") = True
    gUnsafeAutofixRules("footnote_harts") = True
    gUnsafeAutofixRules("footnote_terminal_full_stop") = True
    gUnsafeAutofixRules("footnote_initial_capital") = True
    gUnsafeAutofixRules("footnote_abbreviation") = True
    gUnsafeAutofixRules("footnote_abbreviation_dictionary") = True
    gUnsafeAutofixRules("footnotes_not_endnotes") = True
    gUnsafeAutofixRules("bracket_integrity") = True
    gUnsafeAutofixRules("triplicate_punctuation") = True
    gUnsafeAutofixRules("slash_style") = True
    gUnsafeAutofixRules("custom_rule") = True
    gUnsafeAutofixRules("brand_name_enforcement") = True

    ' -- Tracked-safe allow-list --
    ' Only add a rule here after it has been proven exact-match safe
    ' on large, heavily-redlined documents.
    Set gTrackedSafeRules = CreateObject("Scripting.Dictionary")
    gTrackedSafeRules("double_spaces") = True

    ' -- Comment-safe allow-list (rules that may create inline comments) --
    Set gCommentSafeRules = CreateObject("Scripting.Dictionary")
    gCommentSafeRules("spellchecker") = True
    gCommentSafeRules("licence_license") = True
    gCommentSafeRules("check_cheque") = True
    gCommentSafeRules("repeated_words") = True
    gCommentSafeRules("always_capitalise_terms") = True
    gCommentSafeRules("mandated_legal_term_forms") = True
    gCommentSafeRules("brand_name_enforcement") = True
    gCommentSafeRules("custom_rule") = True
    ' bracket_integrity is intentionally EXCLUDED from comment-safe.
    ' Bracket findings are structural and should be report-only.
    ' Footnote rules are intentionally EXCLUDED from comment-safe.
    ' They default to grouped/report mode, not inline comments.
    gCommentSafeRules("known_anglicised_terms_not_italic") = True
    gCommentSafeRules("foreign_names_not_italic") = True
    gCommentSafeRules("date_time_format") = True
    gCommentSafeRules("currency_number_format") = True
    gCommentSafeRules("spell_out_under_ten") = True
End Sub

' Returns the bucket name for a given rule (for threshold grouping)
Private Function GetRuleBucket(ByVal ruleName As String) As String
    Dim rn As String
    rn = LCase$(ruleName)
    Select Case rn
        Case "spellchecker", "spelling", "licence_license", "check_cheque"
            GetRuleBucket = "spelling"
        Case "footnote_integrity", "footnote_harts", "footnote_terminal_full_stop", _
             "footnote_initial_capital", "footnote_abbreviation", _
             "footnotes_not_endnotes", "footnote_rules", "duplicate_footnotes"
            GetRuleBucket = "footnote"
        Case "double_spaces", "missing_space_after_dot", "space_before_punct", _
             "double_commas", "trailing_spaces"
            GetRuleBucket = "spacing"
        Case "hyphens", "dash_usage"
            GetRuleBucket = "dash"
        Case Else
            GetRuleBucket = ""
    End Select
End Function

' Count findings by bucket in a collection
Private Function CountByBucket(ByVal issues As Collection, ByVal bucket As String) As Long
    Dim cnt As Long
    cnt = 0
    Dim i As Long
    For i = 1 To issues.Count
        If GetRuleBucket(CStr(GetIssueProp(issues(i), "RuleName"))) = bucket Then
            cnt = cnt + 1
        End If
    Next i
    CountByBucket = cnt
End Function

' Build grouped spelling data from issues collection
Private Sub BuildSpellingGroups(ByVal issues As Collection)
    Set gSpellingGroups = CreateObject("Scripting.Dictionary")
    Set gSpellingExamples = CreateObject("Scripting.Dictionary")
    gGroupedSpellingCount = 0
    Dim i As Long
    For i = 1 To issues.Count
        Dim finding As Object
        Set finding = issues(i)
        If GetRuleBucket(CStr(GetIssueProp(finding, "RuleName"))) = "spelling" Then
            gGroupedSpellingCount = gGroupedSpellingCount + 1
            Dim matched As String
            Dim repTxt As String
            matched = CStr(GetIssueProp(finding, "MatchedText"))
            ' Canonical key: MatchedText + ReplacementText (not Suggestion).
            ' ReplacementText is the actual correction; Suggestion is prose.
            ' Fall back to Suggestion only when ReplacementText is absent.
            repTxt = ""
            If HasReplacementText(finding) Then
                repTxt = CStr(GetIssueProp(finding, "ReplacementText"))
            End If
            If Len(repTxt) = 0 Then
                repTxt = CStr(GetIssueProp(finding, "Suggestion"))
            End If
            If Len(matched) = 0 Then matched = "(unknown)"
            If Len(repTxt) = 0 Then repTxt = "(no replacement)"
            Dim pairKey As String
            pairKey = LCase$(matched) & "|" & LCase$(repTxt)
            If gSpellingGroups.Exists(pairKey) Then
                gSpellingGroups(pairKey) = CLng(gSpellingGroups(pairKey)) + 1
            Else
                gSpellingGroups(pairKey) = 1
            End If
            ' Store first 3 example locations
            If Not gSpellingExamples.Exists(pairKey) Then
                gSpellingExamples(pairKey) = CStr(GetIssueProp(finding, "Location"))
            Else
                Dim existing As String
                existing = CStr(gSpellingExamples(pairKey))
                ' Count pipes to see how many we have
                Dim pipeCount As Long
                Dim pc As Long
                pipeCount = 0
                For pc = 1 To Len(existing)
                    If Mid$(existing, pc, 1) = "|" Then pipeCount = pipeCount + 1
                Next pc
                If pipeCount < 2 Then
                    gSpellingExamples(pairKey) = existing & "|" & CStr(GetIssueProp(finding, "Location"))
                End If
            End If
        End If
    Next i
End Sub

' Build grouped footnote data from issues collection
Private Sub BuildFootnoteGroups(ByVal issues As Collection)
    Set gFootnoteGroups = CreateObject("Scripting.Dictionary")
    gGroupedFootnoteCount = 0
    Dim i As Long
    For i = 1 To issues.Count
        Dim finding As Object
        Set finding = issues(i)
        If GetRuleBucket(CStr(GetIssueProp(finding, "RuleName"))) = "footnote" Then
            gGroupedFootnoteCount = gGroupedFootnoteCount + 1
            Dim issText As String
            issText = CStr(GetIssueProp(finding, "Issue"))
            Dim loc As String
            loc = CStr(GetIssueProp(finding, "Location"))
            Dim fnKey As String
            fnKey = LCase$(Left$(issText, 80))
            If gFootnoteGroups.Exists(fnKey) Then
                Dim grp As Object
                Set grp = gFootnoteGroups(fnKey)
                grp("Count") = CLng(grp("Count")) + 1
                If Len(CStr(grp("Locations"))) < 500 Then
                    grp("Locations") = CStr(grp("Locations")) & "|" & loc
                End If
            Else
                Dim newGrp As Object
                Set newGrp = CreateObject("Scripting.Dictionary")
                newGrp("Count") = 1
                newGrp("IssueText") = issText
                newGrp("Locations") = loc
                Set gFootnoteGroups(fnKey) = newGrp
            End If
        End If
    Next i
End Sub

' Check if a rule is in an unsafe autofix category
Public Function IsUnsafeAutofixRule(ByVal ruleName As String) As Boolean
    If gUnsafeAutofixRules Is Nothing Then
        IsUnsafeAutofixRule = False
        Exit Function
    End If
    IsUnsafeAutofixRule = gUnsafeAutofixRules.Exists(LCase$(ruleName))
End Function

' Strong anchor check for tracked changes
Public Function IsStrongTrackedAnchor(ByVal matchedText As String) As Boolean
    IsStrongTrackedAnchor = False
    If Len(matchedText) = 0 Then Exit Function
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(matchedText)
        ch = Mid$(matchedText, i, 1)
        If (ch >= "A" And ch <= "Z") Or _
           (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Then
            IsStrongTrackedAnchor = True
            Exit Function
        End If
    Next i
End Function

' ============================================================
'  SECTION A: CENTRALISED FINDING OUTPUT MODE CLASSIFICATION
'  Decides how a finding should be output based on rule name,
'  bucket, document complexity, thresholds, and allow-lists.
' ============================================================

' Returns the output mode for a finding.
' Possible return values: OUTPUT_TRACKED_SAFE, OUTPUT_COMMENT_ONLY,
'   OUTPUT_REPORT_ONLY, OUTPUT_GROUPED_REPORT
Public Function GetFindingOutputMode(ByVal finding As Object, _
                                      Optional doc As Document = Nothing) As String
    Dim rn As String
    rn = LCase$(CStr(GetIssueProp(finding, "RuleName")))
    Dim bucket As String
    bucket = GetRuleBucket(rn)
    Dim autoFix As Boolean
    On Error Resume Next
    autoFix = CBool(GetIssueProp(finding, "AutoFixSafe"))
    If Err.Number <> 0 Then autoFix = False: Err.Clear
    On Error GoTo 0

    ' Step 0: Hard-block -- these rules must NEVER be tracked-safe,
    ' regardless of AutoFixSafe, allow-list contents, or any other gate.
    ' Defence-in-depth: even if gTrackedSafeRules is accidentally populated,
    ' these rules cannot become OUTPUT_TRACKED_SAFE.
    Dim hardBlock As Boolean
    hardBlock = False
    Select Case rn
        ' Spelling / legal terms
        Case "spellchecker", "licence_license", "check_cheque", _
             "repeated_words", "always_capitalise_terms", _
             "mandated_legal_term_forms"
            hardBlock = True
        ' Structural / punctuation / spacing
        Case "hyphens", "dash_usage", "bracket_integrity", _
             "slash_style", "triplicate_punctuation", _
             "missing_space_after_dot", _
             "space_before_punct", "double_commas", "trailing_spaces"
            hardBlock = True
        ' Footnote rules
        Case "footnote_integrity", "footnote_harts", _
             "footnote_terminal_full_stop", "footnote_initial_capital", _
             "footnote_abbreviation", "footnote_abbreviation_dictionary", _
             "footnotes_not_endnotes", "footnote_rules", "duplicate_footnotes"
            hardBlock = True
        ' Brand / custom rules
        Case "brand_name_enforcement", "custom_rule"
            hardBlock = True
    End Select

    ' Step 1: Check if this should be grouped report
    If ShouldForceGroupedReport(bucket, GetBucketCount(bucket), doc) Then
        GetFindingOutputMode = OUTPUT_GROUPED_REPORT
        Exit Function
    End If

    ' Step 2: Check if rule is explicitly tracked-safe AND has replacement
    ' Skip entirely for hard-blocked rules.
    If (Not hardBlock) And autoFix And IsTrackedSafeRule(rn) And HasReplacementText(finding) Then
        ' Additional gate: document must not be complex for inline markup
        If Not DocumentLooksComplexForInlineMarkup(doc) Then
            ' Check operation type is allowed
            Dim origText As String, repText As String
            origText = CStr(GetIssueProp(finding, "MatchedText"))
            repText = CStr(GetIssueProp(finding, "ReplacementText"))
            Dim opType As String
            opType = GetReplacementOperationType(origText, repText)
            If IsOperationTypeAllowed(rn, opType) Then
                GetFindingOutputMode = OUTPUT_TRACKED_SAFE
                Exit Function
            End If
        End If
    End If

    ' Step 3: Check if rule is comment-safe
    If IsCommentSafeRule(rn) Then
        GetFindingOutputMode = OUTPUT_COMMENT_ONLY
        Exit Function
    End If

    ' Default: report-only
    GetFindingOutputMode = OUTPUT_REPORT_ONLY
End Function

' Is this rule explicitly in the tracked-safe allow-list?
Public Function IsTrackedSafeRule(ByVal ruleName As String) As Boolean
    If gTrackedSafeRules Is Nothing Then EnsureAllowListsInitialized
    IsTrackedSafeRule = gTrackedSafeRules.Exists(LCase$(ruleName))
End Function

' Lazy-init allow-list dictionaries so public queries work
' before a full engine run (e.g. from unit tests).
Private Sub EnsureAllowListsInitialized()
    If Not gTrackedSafeRules Is Nothing Then Exit Sub
    InitGroupedReportState
End Sub

' Is this rule allowed to create inline comments?
Public Function IsCommentSafeRule(ByVal ruleName As String) As Boolean
    If gCommentSafeRules Is Nothing Then
        ' Before initialisation, default to safe (no comments).
        ' This prevents accidental inline comments if a caller
        ' invokes comment logic before InitGroupedReportState.
        IsCommentSafeRule = False
        Exit Function
    End If
    IsCommentSafeRule = gCommentSafeRules.Exists(LCase$(ruleName))
End Function

' Should findings in this bucket be forced to grouped report mode?
Public Function ShouldForceGroupedReport(ByVal bucket As String, _
                                          ByVal totalCount As Long, _
                                          Optional doc As Document = Nothing) As Boolean
    ShouldForceGroupedReport = False
    Select Case bucket
        Case "spelling"
            If totalCount > SPELLING_COMMENT_THRESHOLD Then
                ShouldForceGroupedReport = True
            End If
        Case "footnote"
            If totalCount > FOOTNOTE_COMMENT_THRESHOLD Then
                ShouldForceGroupedReport = True
            End If
        Case "spacing"
            ' Spacing findings are always report-only (no inline comments)
            ShouldForceGroupedReport = True
        Case "dash"
            ' Dash findings are always report-only
            ShouldForceGroupedReport = True
    End Select
    ' Complex documents bias towards grouped report
    If Not ShouldForceGroupedReport And gDocIsComplex Then
        If totalCount > 10 Then ShouldForceGroupedReport = True
    End If
End Function

' Assess whether the document is complex enough that inline markup is risky.
' Factors: existing tracked changes, many comments, Simple Markup view,
' footnote count, revision count.
Public Function DocumentLooksComplexForInlineMarkup( _
        Optional doc As Document = Nothing) As Boolean
    ' Use cached result if already computed
    DocumentLooksComplexForInlineMarkup = gDocIsComplex
End Function

' Compute document complexity (called once per run after doc is available)
Private Sub ComputeDocumentComplexity(doc As Document)
    gDocIsComplex = False
    If doc Is Nothing Then Exit Sub
    On Error Resume Next
    ' Check revision count
    Dim revCount As Long
    revCount = doc.Revisions.Count
    If Err.Number <> 0 Then revCount = 0: Err.Clear
    If revCount > 20 Then gDocIsComplex = True
    ' Check comment count
    Dim cmtCount As Long
    cmtCount = doc.Comments.Count
    If Err.Number <> 0 Then cmtCount = 0: Err.Clear
    If cmtCount > 30 Then gDocIsComplex = True
    ' Check footnote count
    Dim fnCount As Long
    fnCount = doc.Footnotes.Count
    If Err.Number <> 0 Then fnCount = 0: Err.Clear
    If fnCount > 100 Then gDocIsComplex = True
    ' Check document length (very long = complex)
    Dim docLen As Long
    docLen = doc.Content.End
    If Err.Number <> 0 Then docLen = 0: Err.Clear
    If docLen > 200000 Then gDocIsComplex = True
    On Error GoTo 0
    If gDocIsComplex Then
        DebugLog "DocumentComplexity: COMPLEX (revisions=" & revCount & _
                 " comments=" & cmtCount & " footnotes=" & fnCount & _
                 " length=" & docLen & ")"
    End If
End Sub

' Get count of findings in a bucket from current grouped state
Private Function GetBucketCount(ByVal bucket As String) As Long
    Select Case bucket
        Case "spelling": GetBucketCount = gGroupedSpellingCount
        Case "footnote": GetBucketCount = gGroupedFootnoteCount
        Case Else: GetBucketCount = 0
    End Select
End Function

' ============================================================
'  SECTION B: REPLACEMENT OPERATION TYPE CLASSIFICATION
' ============================================================

' Classify what kind of text operation a replacement represents.
Public Function GetReplacementOperationType(ByVal origText As String, _
                                             ByVal newText As String) As String
    If Len(origText) = 0 And Len(newText) > 0 Then
        GetReplacementOperationType = "INSERT"
        Exit Function
    End If
    If Len(newText) = 0 And Len(origText) > 0 Then
        GetReplacementOperationType = "DELETE"
        Exit Function
    End If
    If Len(origText) = 0 And Len(newText) = 0 Then
        GetReplacementOperationType = "UNKNOWN"
        Exit Function
    End If
    ' Check if only whitespace changed
    Dim origStripped As String, newStripped As String
    origStripped = StripAllWhitespace(origText)
    newStripped = StripAllWhitespace(newText)
    If origStripped = newStripped Then
        GetReplacementOperationType = "WHITESPACE_NORMALISE"
        Exit Function
    End If
    ' Check if only punctuation changed (e.g. dash replacement)
    If IsPunctuationOnlyChange(origText, newText) Then
        GetReplacementOperationType = "PUNCTUATION_NORMALISE"
        Exit Function
    End If
    ' General replacement
    GetReplacementOperationType = "REPLACE"
End Function

' Is this operation type allowed for this rule?
Private Function IsOperationTypeAllowed(ByVal ruleName As String, _
                                         ByVal opType As String) As Boolean
    ' Only allow REPLACE for tracked-safe rules (the normal case: word swap)
    ' WHITESPACE_NORMALISE is safe.  PUNCTUATION_NORMALISE is safe for
    ' some rules but NOT for dash/hyphen rules.
    Select Case opType
        Case "REPLACE", "WHITESPACE_NORMALISE"
            IsOperationTypeAllowed = True
        Case "PUNCTUATION_NORMALISE"
            ' Only allow for non-dash rules
            Dim rn As String
            rn = LCase$(ruleName)
            If rn = "hyphens" Or rn = "dash_usage" Then
                IsOperationTypeAllowed = False
            Else
                IsOperationTypeAllowed = True
            End If
        Case "DELETE"
            ' Never allow auto-delete via tracked changes
            IsOperationTypeAllowed = False
        Case "INSERT"
            IsOperationTypeAllowed = False
        Case Else
            IsOperationTypeAllowed = False
    End Select
End Function

' Strip all whitespace from a string (for comparison)
Private Function StripAllWhitespace(ByVal s As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch <> " " And ch <> vbTab And ch <> ChrW(160) And _
           ch <> vbCr And ch <> vbLf And ch <> Chr(11) Then
            result = result & ch
        End If
    Next i
    StripAllWhitespace = result
End Function

' Check if the change is punctuation-only (no alphanumeric chars changed)
Private Function IsPunctuationOnlyChange(ByVal orig As String, _
                                          ByVal repl As String) As Boolean
    ' Extract only alphanumeric chars from both; if equal, it's punct-only
    Dim origAlpha As String, replAlpha As String
    Dim i As Long, ch As String, c As Long
    For i = 1 To Len(orig)
        c = AscW(Mid$(orig, i, 1))
        If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Then
            origAlpha = origAlpha & Mid$(orig, i, 1)
        End If
    Next i
    For i = 1 To Len(repl)
        c = AscW(Mid$(repl, i, 1))
        If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) Then
            replAlpha = replAlpha & Mid$(repl, i, 1)
        End If
    Next i
    ' Both strings are non-empty (caller guarantees this).  If the alpha
    ' content is identical (including both-empty = pure punctuation), the
    ' change is punctuation-only.  E.g. "-" -> en-dash.
    IsPunctuationOnlyChange = (origAlpha = replAlpha)
End Function

' Public accessors for grouped report data
Public Function GetGroupedSpellingCount() As Long
    GetGroupedSpellingCount = gGroupedSpellingCount
End Function

Public Function GetGroupedFootnoteCount() As Long
    GetGroupedFootnoteCount = gGroupedFootnoteCount
End Function

Public Function GetCommentsCreated() As Long
    GetCommentsCreated = gCommentsCreated
End Function

Public Function GetTrackedEditsApplied() As Long
    GetTrackedEditsApplied = gTrackedEditsApplied
End Function

Public Function GetEditsSkippedUnsafe() As Long
    GetEditsSkippedUnsafe = gEditsSkippedUnsafe
End Function

Public Function GetSpellingGroupsDict() As Object
    Set GetSpellingGroupsDict = gSpellingGroups
End Function

Public Function GetSpellingExamplesDict() As Object
    Set GetSpellingExamplesDict = gSpellingExamples
End Function

Public Function GetFootnoteGroupsDict() As Object
    Set GetFootnoteGroupsDict = gFootnoteGroups
End Function

' Generate plain-text report
Public Function GenerateTextReport(issues As Collection, _
                                    filePath As String, _
                                    Optional doc As Document = Nothing) As String
    TraceEnter "GenerateTextReport"
    Dim fileNum As Integer
    Dim i As Long

    ' Resolve document name
    Dim docName As String
    On Error Resume Next
    If Not doc Is Nothing Then
        docName = doc.Name
    Else
        docName = "(no document)"
    End If
    If Err.Number <> 0 Then docName = "(unknown)": Err.Clear
    On Error GoTo 0

    fileNum = FreeFile
    On Error Resume Next
    Open filePath For Output As #fileNum
    If Err.Number <> 0 Then
        GenerateTextReport = "Error: could not write to " & filePath
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    On Error GoTo TxtReportErr

    ' ASCII art header
    Print #fileNum, ""
    Print #fileNum, " .==================-.   :=======.        -=======."
    Print #fileNum, " .:::::::::===========-  =========       ========="
    Print #fileNum, "         =============== .=========.     =========-"
    Print #fileNum, "  :::::::......:======== :=========-   :==========."
    Print #fileNum, " -======.        ======: ===========. -==========="
    Print #fileNum, " .=======        :======..===========--===========-"
    Print #fileNum, " :=======      .======== :=================:======."
    Print #fileNum, " -=====================  ======: ========- ======="
    Print #fileNum, " .====================   .======  :======:  ======-"
    Print #fileNum, " :======-  =======-      -======   -====   :======."
    Print #fileNum, " -======    -=======     ======:    ==-    ======="
    Print #fileNum, "  =======     -=======   :======     .:     ======:"
    Print #fileNum, " :=======      :=======  -======           :======."
    Print #fileNum, " -======.       :======= ======:           ======="
    Print #fileNum, ""
    Print #fileNum, "  ::: ::  -:-:- :-:.--::-:--  -:.:: ::  :-:-:..:-.::.:-"
    Print #fileNum, "  .:.:- =  = = =--:-=:.=. =.=  - = == - :::-:-=.=  = -.-:"
    Print #fileNum, ""
    Print #fileNum, "  PLEADINGS CHECKER REPORT"
    Print #fileNum, "============================================================"
    Print #fileNum, ""
    Print #fileNum, "Document:  " & docName
    Print #fileNum, "Date:      " & Format(Now, "yyyy-mm-dd hh:nn:ss")

    Dim prStr As String
    prStr = pageRangeString
    If Len(prStr) = 0 Then prStr = "all"
    Print #fileNum, "Pages:     " & prStr
    Print #fileNum, "Issues:    " & issues.Count
    Print #fileNum, ""

    ' -- Grouped Spelling Section --
    If gGroupedSpellingCount > 0 Then
        Print #fileNum, "============================================================"
        Print #fileNum, "  SPELLING (" & gGroupedSpellingCount & " occurrences, " & _
                        gSpellingGroups.Count & " unique pairs)"
        Print #fileNum, "============================================================"
        Dim spKeys As Variant
        spKeys = gSpellingGroups.keys
        Dim sk As Long
        For sk = 0 To gSpellingGroups.Count - 1
            Dim pairParts() As String
            pairParts = Split(CStr(spKeys(sk)), "|")
            Dim spWrong As String, spRight As String
            spWrong = pairParts(0)
            If UBound(pairParts) >= 1 Then spRight = pairParts(1) Else spRight = "?"
            Print #fileNum, "  " & spWrong & " -> " & spRight & _
                            "  (" & gSpellingGroups(spKeys(sk)) & " occurrences)"
            If gSpellingExamples.Exists(CStr(spKeys(sk))) Then
                Dim exLocs() As String
                exLocs = Split(CStr(gSpellingExamples(spKeys(sk))), "|")
                Dim ex As Long
                For ex = 0 To UBound(exLocs)
                    Print #fileNum, "    - " & exLocs(ex)
                Next ex
            End If
        Next sk
        Print #fileNum, ""
    End If

    ' -- Grouped Footnote Section --
    If gGroupedFootnoteCount > 0 Then
        Print #fileNum, "============================================================"
        Print #fileNum, "  FOOTNOTES (" & gGroupedFootnoteCount & " findings)"
        Print #fileNum, "============================================================"
        Dim fnKeys As Variant
        fnKeys = gFootnoteGroups.keys
        Dim fk As Long
        For fk = 0 To gFootnoteGroups.Count - 1
            Dim fnGrp As Object
            Set fnGrp = gFootnoteGroups(fnKeys(fk))
            Print #fileNum, "  " & CStr(fnGrp("IssueText"))
            Print #fileNum, "    Count: " & CStr(fnGrp("Count"))
            Dim fnLocs() As String
            fnLocs = Split(CStr(fnGrp("Locations")), "|")
            Dim fl As Long
            For fl = 0 To UBound(fnLocs)
                If fl < 5 Then Print #fileNum, "    - " & fnLocs(fl)
            Next fl
            If UBound(fnLocs) >= 5 Then
                Print #fileNum, "    ... and " & (UBound(fnLocs) - 4) & " more"
            End If
        Next fk
        Print #fileNum, ""
    End If

    ' -- By UI label --
    Print #fileNum, "============================================================"
    Print #fileNum, "  ALL FINDINGS BY TYPE"
    Print #fileNum, "============================================================"

    ' Group by UI label
    Dim uiGroups As Object
    Set uiGroups = CreateObject("Scripting.Dictionary")
    For i = 1 To issues.Count
        Dim uiLbl As String
        uiLbl = GetUILabel(CStr(GetIssueProp(issues(i), "RuleName")))
        If Not uiGroups.Exists(uiLbl) Then
            Dim grpColl As Collection
            Set grpColl = New Collection
            Set uiGroups(uiLbl) = grpColl
        End If
        Dim theGroup As Collection
        Set theGroup = uiGroups(uiLbl)
        theGroup.Add issues(i)
    Next i

    Dim uiKeys As Variant
    uiKeys = uiGroups.keys
    Dim uk As Long
    For uk = 0 To uiGroups.Count - 1
        Dim labelGroup As Collection
        Set labelGroup = uiGroups(uiKeys(uk))
        Print #fileNum, ""
        Print #fileNum, "--- " & CStr(uiKeys(uk)) & " (" & labelGroup.Count & ") ---"
        Dim gi As Long
        Dim maxItems As Long
        maxItems = labelGroup.Count
        If maxItems > 50 Then maxItems = 50
        For gi = 1 To maxItems
            Dim gFinding As Object
            Set gFinding = labelGroup(gi)
            Print #fileNum, "  [" & CStr(GetIssueProp(gFinding, "Location")) & "] " & _
                            CStr(GetIssueProp(gFinding, "Issue"))
        Next gi
        If labelGroup.Count > 50 Then
            Print #fileNum, "  ... and " & (labelGroup.Count - 50) & " more"
        End If
    Next uk

    Print #fileNum, ""
    Print #fileNum, "============================================================"
    Print #fileNum, "  END OF REPORT"
    Print #fileNum, "============================================================"

    Close #fileNum
    GenerateTextReport = "Text report saved: " & filePath
    TraceExit "GenerateTextReport"
    Exit Function

TxtReportErr:
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    GenerateTextReport = "Error writing text report: Err " & Err.Number & ": " & Err.Description
    TraceExit "GenerateTextReport", "FAILED"
End Function

' Print debug summary to Immediate window
Public Sub PrintDebugSummary(ByVal issues As Collection)
    Debug.Print "=== PLEADINGS CHECKER DEBUG SUMMARY ==="
    Debug.Print "  Total issues:           " & issues.Count
    Debug.Print "  Grouped spelling count:  " & gGroupedSpellingCount
    Debug.Print "  Grouped footnote count:  " & gGroupedFootnoteCount
    Debug.Print "  Comments created:        " & gCommentsCreated
    Debug.Print "  Tracked edits applied:   " & gTrackedEditsApplied
    Debug.Print "  Edits skipped (unsafe):  " & gEditsSkippedUnsafe
    Debug.Print "  Top 5 slowest: " & GetTopSlowestRules(5)
    Debug.Print "=== END DEBUG SUMMARY ==="
End Sub

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

    ' Multiple candidates: show picker with browse option
    Dim prompt As String
    prompt = "Select the document to check:" & vbCrLf & vbCrLf
    Dim idx As Long
    For idx = 1 To candidates.Count
        prompt = prompt & idx & ". " & candidates(idx).Name & vbCrLf
    Next idx
    prompt = prompt & vbCrLf & "Or type B to browse for a file."

    Dim selectionText As String
    selectionText = InputBox(prompt, "Pleadings Checker - Select Document", "1")
    If Len(Trim(selectionText)) = 0 Then Exit Function  ' cancelled

    ' Browse option
    If UCase$(Trim$(selectionText)) = "B" Then
        Set GetTargetDocument = BrowseForTargetDocument()
        Exit Function
    End If

    Dim chosen As Long
    If IsNumeric(selectionText) Then
        chosen = CLng(selectionText)
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
'  BROWSE FOR TARGET DOCUMENT (Section H)
'  Opens a file dialog to select and open a Word document.
'  Falls back to InputBox path entry on Mac or if FileDialog
'  is not available.
' ============================================================
Private Function BrowseForTargetDocument() As Document
    Set BrowseForTargetDocument = Nothing
    Dim filePath As String
    filePath = ""

    ' Try FileDialog first (Windows)
    On Error Resume Next
    Dim fd As Object
    Set fd = Application.FileDialog(1)  ' msoFileDialogOpen
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ' Fallback: InputBox for path
        filePath = InputBox("Enter the full path to the document:", _
                            "Pleadings Checker - Open Document", "")
        If Len(Trim$(filePath)) = 0 Then Exit Function
    Else
        On Error GoTo 0
        With fd
            .Title = "Select Document to Check"
            .AllowMultiSelect = False
            On Error Resume Next
            .Filters.Clear
            .Filters.Add "Word Documents", "*.docx;*.doc;*.docm"
            .Filters.Add "All Files", "*.*"
            If Err.Number <> 0 Then Err.Clear
            On Error GoTo 0
            If .Show = -1 Then
                filePath = CStr(.SelectedItems(1))
            Else
                Exit Function  ' cancelled
            End If
        End With
    End If

    ' Open the selected file
    If Len(filePath) > 0 Then
        On Error Resume Next
        Dim openDoc As Document
        Set openDoc = Documents.Open(filePath)
        If Err.Number <> 0 Then
            MsgBox "Could not open:" & vbCrLf & filePath & vbCrLf & _
                   "Error: " & Err.Description, vbExclamation, "Pleadings Checker"
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
        Set BrowseForTargetDocument = openDoc
    End If
End Function

' ============================================================
'  QUICK RUN (fallback when launcher/form is not imported)
'  Prompts for target document via GetTargetDocument(), then
'  asks for an optional page range before running all rules
'  with UK spelling.  Report/review only -- does NOT auto-edit
'  the document.  Use the full form for apply options.
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

    ' Prompt for page range (blank = all pages)
    Dim pgInput As String
    pgInput = InputBox("Page range (e.g. 1,3,5-8) or blank for all pages:", _
                        "Pleadings Checker - Page Range", "")
    SetPageRangeFromString Trim(pgInput)

    SetSpellingMode "UK"

    Dim issues As Collection
    Set issues = RunAllPleadingsRules(targetDoc, cfg)

    Dim summary As String
    summary = GetIssueSummary(issues)

    If issues.Count = 0 Then
        MsgBox "No issues found.", vbInformation, "Pleadings Checker"
    Else
        MsgBox summary & vbCrLf & vbCrLf & _
               "Use the full form (PleadingsChecker) to apply suggestions.", _
               vbInformation, "Pleadings Checker"
    End If
    TraceExit "RunQuick", issues.Count & " issues"
End Sub

' ============================================================
'  RunCheckerFromFormConfig (Section G)
'  Single entry point for the form to run all checks.
'  Accepts a config dictionary with all UI state pre-gathered
'  so the form does not need to call individual Set* methods.
'
'  Config keys (all optional, with defaults):
'    "ruleConfig"      -> Dictionary of rule_name -> Boolean
'    "pageRange"       -> String (page range spec)
'    "spellingMode"    -> "UK" or "US"
'    "quoteNesting"    -> "SINGLE" or "DOUBLE"
'    "smartQuotePref"  -> "SMART" or "STRAIGHT"
'    "dateFormatPref"  -> "UK" or "US"
'    "termFormatPref"  -> "BOLD" / "BOLDITALIC" / "ITALIC" / "NONE"
'    "termQuotePref"   -> "SINGLE" or "DOUBLE"
'    "spaceStylePref"  -> "ONE" or "TWO"
'    "nonEngTermPref"  -> "ITALICS" or "REGULAR"
' ============================================================
Public Function RunCheckerFromFormConfig(doc As Document, _
                                          formConfig As Object) As Collection
    TraceEnter "RunCheckerFromFormConfig"

    ' Apply all preferences from the config dictionary
    If formConfig.Exists("pageRange") Then
        SetPageRangeFromString CStr(formConfig("pageRange"))
    Else
        SetPageRangeFromString ""
    End If
    If formConfig.Exists("spellingMode") Then SetSpellingMode CStr(formConfig("spellingMode"))
    If formConfig.Exists("quoteNesting") Then SetQuoteNesting CStr(formConfig("quoteNesting"))
    If formConfig.Exists("smartQuotePref") Then SetSmartQuotePref CStr(formConfig("smartQuotePref"))
    If formConfig.Exists("dateFormatPref") Then SetDateFormatPref CStr(formConfig("dateFormatPref"))
    If formConfig.Exists("termFormatPref") Then SetTermFormatPref CStr(formConfig("termFormatPref"))
    If formConfig.Exists("termQuotePref") Then SetTermQuotePref CStr(formConfig("termQuotePref"))
    If formConfig.Exists("spaceStylePref") Then SetSpaceStylePref CStr(formConfig("spaceStylePref"))
    If formConfig.Exists("nonEngTermPref") Then SetNonEngTermPref CStr(formConfig("nonEngTermPref"))

    ' Get rule config
    Dim cfg As Object
    If formConfig.Exists("ruleConfig") Then
        Set cfg = formConfig("ruleConfig")
    Else
        Set cfg = InitRuleConfig()
    End If

    ' Run all rules
    Set RunCheckerFromFormConfig = RunAllPleadingsRules(doc, cfg)
    TraceExit "RunCheckerFromFormConfig"
End Function

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
    cfg.Add "duplicate_footnotes", False
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

    ' -- Region cache initialisation --
    Const BODY_TEXT_MIN_LEN As Long = 200
    coverPageEnd = -1

    Dim bqCap As Long: bqCap = 64
    ReDim bqStarts(0 To bqCap - 1)
    ReDim bqEnds(0 To bqCap - 1)
    bqCount = 0

    Dim tocCap As Long: tocCap = 16
    ReDim tocStarts(0 To tocCap - 1)
    ReDim tocEnds(0 To tocCap - 1)
    tocCount = 0

    ' -- Collect TOC field ranges first (separate from paragraph scan) --
    On Error Resume Next
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
    On Error GoTo 0

    ' -- Smart-quote multi-paragraph tracking --
    Dim insideMultiParaQuote As Boolean
    insideMultiParaQuote = False

    ' -- Single pass over all paragraphs --
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        ' Cache paragraph start position
        If paraStartCount >= cap Then
            cap = cap * 2
            ReDim Preserve paraStartPos(0 To cap - 1)
        End If
        Dim pStart As Long, pEnd As Long
        pStart = para.Range.Start
        pEnd = para.Range.End
        If Err.Number <> 0 Then Err.Clear: GoTo NextCachePara
        paraStartPos(paraStartCount) = pStart
        paraStartCount = paraStartCount + 1

        ' -- Cover page detection --
        If coverPageEnd < 0 Then
            Dim cpText As String
            cpText = ""
            cpText = para.Range.Text
            If Err.Number <> 0 Then cpText = "": Err.Clear
            If Len(cpText) > 0 Then
                If Right$(cpText, 1) = vbCr Or Right$(cpText, 1) = Chr(13) Then
                    cpText = Left$(cpText, Len(cpText) - 1)
                End If
            End If
            Dim cleanCpText As String
            cleanCpText = Replace(Replace(Replace(cpText, vbLf, ""), vbVerticalTab, ""), Chr(11), "")
            If Len(cleanCpText) > BODY_TEXT_MIN_LEN Then
                coverPageEnd = pStart
            End If
        End If

        ' -- TOC paragraph detection (style-based and pattern-based) --
        Dim tocSn As String
        tocSn = ""
        tocSn = LCase(para.Style.NameLocal)
        If Err.Number <> 0 Then tocSn = "": Err.Clear
        Dim isTocPara As Boolean
        isTocPara = False
        If InStr(tocSn, "toc") > 0 Or InStr(tocSn, "table of contents") > 0 Or _
           InStr(tocSn, "contents") > 0 Then
            isTocPara = True
        End If
        If Not isTocPara Then
            Dim tocParaText As String
            tocParaText = ""
            tocParaText = para.Range.Text
            If Err.Number <> 0 Then tocParaText = "": Err.Clear
            If Len(tocParaText) > 3 Then
                If tocParaText Like "*[." & vbTab & "][." & vbTab & "]*#" & vbCr Or _
                   tocParaText Like "*[." & vbTab & "][." & vbTab & "]*#" Then
                    isTocPara = True
                End If
            End If
        End If
        If isTocPara Then
            If tocCount >= tocCap Then
                tocCap = tocCap * 2
                ReDim Preserve tocStarts(0 To tocCap - 1)
                ReDim Preserve tocEnds(0 To tocCap - 1)
            End If
            tocStarts(tocCount) = pStart
            tocEnds(tocCount) = pEnd
            tocCount = tocCount + 1
        End If

        ' -- Block-quote detection using canonical IsBlockQuotePara --
        Dim isBQ As Boolean
        isBQ = False

        ' Try canonical IsBlockQuotePara via Application.Run
        Dim bqResult As Boolean
        bqResult = False
        bqResult = Application.Run("Rules_Formatting.IsBlockQuotePara", para)
        If Err.Number <> 0 Then
            ' Module not imported -- fall back to style-name-only check
            Err.Clear
            Dim fallbackSn As String
            fallbackSn = ""
            fallbackSn = LCase(para.Style.NameLocal)
            If Err.Number <> 0 Then fallbackSn = "": Err.Clear
            If InStr(fallbackSn, "quote") > 0 Or InStr(fallbackSn, "block") > 0 Or _
               InStr(fallbackSn, "extract") > 0 Then
                bqResult = True
            End If
        End If
        isBQ = bqResult

        ' -- Multi-paragraph smart-quote detection --
        Dim pText As String
        pText = ""
        pText = para.Range.Text
        If Err.Number <> 0 Then pText = "": Err.Clear
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

NextCachePara:
    Next para
    On Error GoTo 0

    paraCacheValid = True
    PerfTimerEnd "BuildParagraphCache"
    PerfCount "paragraphs_cached", paraStartCount
    PerfCount "block_quote_regions", bqCount
    PerfCount "toc_regions", tocCount
End Sub

Private Function FindParagraphIndex(ByVal pos As Long) As Long
    If Not paraCacheValid Or paraStartCount = 0 Then
        FindParagraphIndex = 0
        Exit Function
    End If

    ' Binary search for paragraph containing this position
    Dim lo As Long, hi As Long, pivot As Long
    lo = 0
    hi = paraStartCount - 1

    Do While lo <= hi
        pivot = (lo + hi) \ 2
        If pivot < paraStartCount - 1 Then
            If paraStartPos(pivot) <= pos And paraStartPos(pivot + 1) > pos Then
                FindParagraphIndex = pivot + 1  ' 1-based
                Exit Function
            ElseIf paraStartPos(pivot) > pos Then
                hi = pivot - 1
            Else
                lo = pivot + 1
            End If
        Else
            ' Last paragraph
            If paraStartPos(pivot) <= pos Then
                FindParagraphIndex = pivot + 1
            Else
                FindParagraphIndex = pivot
            End If
            Exit Function
        End If
    Loop

    FindParagraphIndex = lo + 1  ' 1-based
End Function

' ============================================================
'  CONSOLIDATED PARAGRAPH-LEVEL RULE RUNNER
'  Iterates doc.Paragraphs exactly once and dispatches to all
'  paragraph-level rule handlers that are enabled in the config.
'  Each handler receives the paragraph's Range, text, start
'  position, and list-prefix length, and appends issues to the
'  shared collection.
' ============================================================
Public Sub RunParagraphRules(doc As Document, config As Object, _
                              ByRef allIssues As Collection)
    PerfTimerStart "paragraph_rules_combined"
    TraceStep "RunParagraphRules", "starting consolidated paragraph pass"

    ' -- Build list of enabled handlers --------------------------
    ' Each entry is the fully qualified name for Application.Run
    Dim handlers() As String
    Dim hCount As Long
    hCount = 0
    ReDim handlers(0 To 15)

    ' Repeated words
    If IsRuleEnabled(config, "repeated_words") Then
        handlers(hCount) = "Rules_TextScan.ProcessParagraph_RepeatedWords"
        hCount = hCount + 1
    End If
    ' Spell out under ten
    If IsRuleEnabled(config, "spell_out_under_ten") Then
        handlers(hCount) = "Rules_TextScan.ProcessParagraph_SpellOutUnderTen"
        hCount = hCount + 1
    End If
    ' Double spaces
    If IsRuleEnabled(config, "double_spaces") Then
        handlers(hCount) = "Rules_Spacing.ProcessParagraph_DoubleSpaces"
        hCount = hCount + 1
    End If
    ' Punctuation sub-rules (all under "punctuation" toggle)
    If IsRuleEnabled(config, "punctuation") Then
        handlers(hCount) = "Rules_Punctuation.ProcessParagraph_TriplicatePunctuation"
        hCount = hCount + 1
        handlers(hCount) = "Rules_Punctuation.ProcessParagraph_DashUsage"
        hCount = hCount + 1
        handlers(hCount) = "Rules_Punctuation.ProcessParagraph_BracketIntegrity"
        hCount = hCount + 1
        handlers(hCount) = "Rules_Spacing.ProcessParagraph_DoubleCommas"
        hCount = hCount + 1
        handlers(hCount) = "Rules_Spacing.ProcessParagraph_MissingSpaceAfterDot"
        hCount = hCount + 1
        handlers(hCount) = "Rules_Spacing.ProcessParagraph_SpaceBeforePunct"
        hCount = hCount + 1
    End If
    ' Always capitalise terms
    If IsRuleEnabled(config, "always_capitalise_terms") Then
        handlers(hCount) = "Rules_LegalTerms.ProcessParagraph_AlwaysCapitalise"
        hCount = hCount + 1
    End If
    ' Non-English terms (italics)
    If IsRuleEnabled(config, "non_english_terms") Then
        handlers(hCount) = "Rules_Italics.ProcessParagraph_AnglicisedTerms"
        hCount = hCount + 1
        handlers(hCount) = "Rules_Italics.ProcessParagraph_ForeignNames"
        hCount = hCount + 1
    End If

    If hCount = 0 Then
        PerfTimerEnd "paragraph_rules_combined"
        Exit Sub
    End If

    ' -- Iterate paragraphs once ---------------------------------
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraText As String
    Dim paraStart As Long
    Dim listPrefixLen As Long
    Dim paraIssues As New Collection
    Dim h As Long
    Dim paraCount As Long
    paraCount = 0

    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaRP
        paraStart = paraRange.Start

        ' Page-range filter
        If IsPastPageFilter(paraStart) Then Exit For
        If Not IsInPageRange(paraRange) Then GoTo NextParaRP

        paraText = paraRange.Text
        If Err.Number <> 0 Then Err.Clear: GoTo NextParaRP
        If Len(paraText) = 0 Then GoTo NextParaRP

        ' Calculate list prefix length
        listPrefixLen = 0
        Dim lStr As String
        lStr = ""
        lStr = para.Range.ListFormat.ListString
        If Err.Number <> 0 Then lStr = "": Err.Clear
        If Len(lStr) > 0 And Len(paraText) > Len(lStr) Then
            If Left$(paraText, Len(lStr)) = lStr Then
                listPrefixLen = Len(lStr)
            End If
        End If

        ' Cancellation check every 50 paragraphs
        paraCount = paraCount + 1
        If paraCount Mod 50 = 0 Then
            DoEvents
            If gCancelRun Then
                Err.Raise ERR_RUN_CANCELLED, "RunParagraphRules", "Run cancelled"
            End If
        End If

        ' -- Dispatch to each enabled handler --------------------
        For h = 0 To hCount - 1
            Err.Clear
            Application.Run handlers(h), doc, paraRange, paraText, _
                paraStart, listPrefixLen, paraIssues
            If Err.Number <> 0 Then
                ' Handler not available or errored - skip silently
                If Err.Number <> ERR_RUN_CANCELLED Then
                    DebugLogError "RunParagraphRules", handlers(h), Err.Number, Err.Description
                    Err.Clear
                Else
                    ' Re-raise cancellation
                    Dim cancelErr As Long
                    cancelErr = Err.Number
                    On Error GoTo 0
                    Err.Raise cancelErr, "RunParagraphRules", "Run cancelled"
                End If
            End If
        Next h

NextParaRP:
    Next para
    On Error GoTo 0

    ' -- Merge paragraph issues into master collection -----------
    Dim pi As Long
    For pi = 1 To paraIssues.Count
        allIssues.Add paraIssues(pi)
    Next pi

    PerfTimerEnd "paragraph_rules_combined"
    PerfCount "paragraph_rules_paragraphs", paraCount
    TraceStep "RunParagraphRules", "completed: " & paraCount & " paragraphs, " & _
              paraIssues.Count & " issues from " & hCount & " handlers"
End Sub

' ============================================================
'  CONSOLIDATED FOOTNOTE/ENDNOTE RULE RUNNER
'  Iterates doc.Footnotes exactly once and doc.Endnotes once,
'  running all enabled footnote/endnote checks per note.
'  Replaces 6 separate loops with 2.
' ============================================================
Public Sub RunFootnoteRules(doc As Document, config As Object, _
                             ByRef allIssues As Collection)
    PerfTimerStart "footnote_rules_combined"
    TraceStep "RunFootnoteRules", "starting consolidated footnote pass"

    Dim fnIssues As New Collection
    Dim fnEnabled As Boolean: fnEnabled = IsRuleEnabled(config, "footnote_rules")
    Dim dupEnabled As Boolean: dupEnabled = IsRuleEnabled(config, "duplicate_footnotes")

    ' -- FootnotesNotEndnotes (no-loop check) ----------------------
    If fnEnabled Then
        AddIssuesToCollection fnIssues, _
            TryRunRule("Rules_FootnoteHarts.Check_FootnotesNotEndnotes", doc)
    End If

    ' -- Build Harts handler list ----------------------------------
    Dim hartsHandlers() As String
    Dim hartsCount As Long
    hartsCount = 0
    ReDim hartsHandlers(0 To 3)

    If fnEnabled Then
        hartsHandlers(hartsCount) = "Rules_FootnoteHarts.ProcessFootnote_TerminalFullStop"
        hartsCount = hartsCount + 1
        hartsHandlers(hartsCount) = "Rules_FootnoteHarts.ProcessFootnote_InitialCapital"
        hartsCount = hartsCount + 1
        hartsHandlers(hartsCount) = "Rules_FootnoteHarts.ProcessFootnote_AbbreviationDictionary"
        hartsCount = hartsCount + 1
    End If

    ' -- Initialise Harts module-level caches ----------------------
    If hartsCount > 0 Then
        On Error Resume Next
        Application.Run "Rules_FootnoteHarts.InitFootnoteCaches"
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If

    ' -- Single pass through footnotes -----------------------------
    If doc.Footnotes.Count > 0 Then
        Dim fnExpectedIdx As Long: fnExpectedIdx = 1
        Dim fnDupDict As Object
        If dupEnabled Then
            Set fnDupDict = CreateObject("Scripting.Dictionary")
        End If

        Dim fi As Long
        Dim fn As Footnote
        Dim fnNoteText As String
        Dim fnRefStart As Long
        Dim fnCharBefore As String
        Dim fnRngBefore As Range
        Dim fnCleanText As String

        For fi = 1 To doc.Footnotes.Count
            On Error Resume Next
            Set fn = doc.Footnotes(fi)
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextFnCombined
            On Error GoTo 0

            On Error Resume Next
            If Not TextAnchoring.IsInPageRange(fn.Reference) Then
                fnExpectedIdx = fnExpectedIdx + 1
                On Error GoTo 0
                GoTo NextFnCombined
            End If
            On Error GoTo 0

            ' -- Integrity: sequence ---------------------------------
            If fnEnabled Then
                If fn.Index <> fnExpectedIdx Then
                    TextAnchoring.AddIssue fnIssues, "footnote_integrity", doc, fn.Reference, _
                        "Footnote numbering gap: expected " & fnExpectedIdx & ", found " & fn.Index, _
                        "Renumber footnotes sequentially", _
                        fn.Reference.Start, fn.Reference.End
                End If
            End If
            fnExpectedIdx = fnExpectedIdx + 1

            ' -- Integrity: placement --------------------------------
            If fnEnabled Then
                fnRefStart = fn.Reference.Start
                If fnRefStart > 0 Then
                    Set fnRngBefore = TextAnchoring.SafeRange(doc, fnRefStart - 1, fnRefStart)
                    If Not fnRngBefore Is Nothing Then
                        fnCharBefore = fnRngBefore.Text
                        If Not TextAnchoring.IsPunctuation(fnCharBefore) Then
                            TextAnchoring.AddIssue fnIssues, "footnote_integrity", doc, fn.Reference, _
                                "Footnote " & fn.Index & " reference not placed after punctuation", _
                                "Place footnote reference after punctuation mark", _
                                fn.Reference.Start, fn.Reference.End
                        End If
                    End If
                End If
            End If

            ' -- Read note text (shared by empty, duplicate, Harts) --
            On Error Resume Next
            fnNoteText = fn.Range.Text
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextFnCombined
            On Error GoTo 0

            ' -- Integrity: empty ------------------------------------
            If fnEnabled Then
                fnCleanText = Trim(Replace(Replace(fnNoteText, vbCr, ""), vbLf, ""))
                If Len(fnCleanText) = 0 Then
                    TextAnchoring.AddIssue fnIssues, "footnote_integrity", doc, fn.Reference, _
                        "Footnote " & fn.Index & " has empty content", _
                        "Add content or remove the empty footnote", _
                        fn.Reference.Start, fn.Reference.End
                End If
            End If

            ' -- Duplicate: build dict / flag duplicates -------------
            If dupEnabled Then
                fnCleanText = Trim(Replace(Replace(fnNoteText, vbCr, ""), vbLf, ""))
                If Len(fnCleanText) > 0 Then
                    If fnDupDict.Exists(fnCleanText) Then
                        Dim fnFirstIdx As Long
                        fnFirstIdx = CLng(fnDupDict(fnCleanText))
                        TextAnchoring.AddIssue fnIssues, "footnote_integrity", doc, fn.Reference, _
                            "Footnote " & fn.Index & " has identical content to footnote " & fnFirstIdx, _
                            "Remove duplicate or differentiate content", _
                            fn.Reference.Start, fn.Reference.End, "possible_error"
                    Else
                        fnDupDict.Add fnCleanText, fn.Index
                    End If
                End If
            End If

            ' -- Harts handlers (terminal stop, initial capital, abbreviation dict) --
            Dim fh As Long
            For fh = 0 To hartsCount - 1
                On Error Resume Next
                Application.Run hartsHandlers(fh), doc, fn, fnNoteText, fnIssues
                If Err.Number <> 0 Then
                    If Err.Number = ERR_RUN_CANCELLED Then
                        On Error GoTo 0
                        Err.Raise ERR_RUN_CANCELLED, "RunFootnoteRules", "Run cancelled"
                    End If
                    DebugLogError "RunFootnoteRules", hartsHandlers(fh), Err.Number, Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            Next fh

            ' -- Cancellation check every 50 notes -------------------
            If fi Mod 50 = 0 Then
                DoEvents
                If gCancelRun Then
                    Err.Raise ERR_RUN_CANCELLED, "RunFootnoteRules", "Run cancelled"
                End If
            End If

NextFnCombined:
        Next fi
    End If

    ' -- Single pass through endnotes (integrity only) -------------
    If doc.Endnotes.Count > 0 And (fnEnabled Or dupEnabled) Then
        Dim enExpectedIdx As Long: enExpectedIdx = 1
        Dim enDupDict As Object
        If dupEnabled Then
            Set enDupDict = CreateObject("Scripting.Dictionary")
        End If

        Dim ei As Long
        Dim en As Endnote
        Dim enNoteText As String
        Dim enRefStart As Long
        Dim enCharBefore As String
        Dim enRngBefore As Range
        Dim enCleanText As String

        For ei = 1 To doc.Endnotes.Count
            On Error Resume Next
            Set en = doc.Endnotes(ei)
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextEnCombined
            On Error GoTo 0

            On Error Resume Next
            If Not TextAnchoring.IsInPageRange(en.Reference) Then
                enExpectedIdx = enExpectedIdx + 1
                On Error GoTo 0
                GoTo NextEnCombined
            End If
            On Error GoTo 0

            ' -- Integrity: sequence ---------------------------------
            If fnEnabled Then
                If en.Index <> enExpectedIdx Then
                    TextAnchoring.AddIssue fnIssues, "footnote_integrity", doc, en.Reference, _
                        "Endnote numbering gap: expected " & enExpectedIdx & ", found " & en.Index, _
                        "Renumber endnotes sequentially", _
                        en.Reference.Start, en.Reference.End
                End If
            End If
            enExpectedIdx = enExpectedIdx + 1

            ' -- Integrity: placement --------------------------------
            If fnEnabled Then
                enRefStart = en.Reference.Start
                If enRefStart > 0 Then
                    Set enRngBefore = TextAnchoring.SafeRange(doc, enRefStart - 1, enRefStart)
                    If Not enRngBefore Is Nothing Then
                        enCharBefore = enRngBefore.Text
                        If Not TextAnchoring.IsPunctuation(enCharBefore) Then
                            TextAnchoring.AddIssue fnIssues, "footnote_integrity", doc, en.Reference, _
                                "Endnote " & en.Index & " reference not placed after punctuation", _
                                "Place endnote reference after punctuation mark", _
                                en.Reference.Start, en.Reference.End
                        End If
                    End If
                End If
            End If

            ' -- Read note text (shared by empty + duplicate) --------
            On Error Resume Next
            enNoteText = en.Range.Text
            If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: GoTo NextEnCombined
            On Error GoTo 0

            ' -- Integrity: empty ------------------------------------
            If fnEnabled Then
                enCleanText = Trim(Replace(Replace(enNoteText, vbCr, ""), vbLf, ""))
                If Len(enCleanText) = 0 Then
                    TextAnchoring.AddIssue fnIssues, "footnote_integrity", doc, en.Reference, _
                        "Endnote " & en.Index & " has empty content", _
                        "Add content or remove the empty endnote", _
                        en.Reference.Start, en.Reference.End
                End If
            End If

            ' -- Duplicate: build dict / flag duplicates -------------
            If dupEnabled Then
                enCleanText = Trim(Replace(Replace(enNoteText, vbCr, ""), vbLf, ""))
                If Len(enCleanText) > 0 Then
                    If enDupDict.Exists(enCleanText) Then
                        Dim enFirstIdx As Long
                        enFirstIdx = CLng(enDupDict(enCleanText))
                        TextAnchoring.AddIssue fnIssues, "footnote_integrity", doc, en.Reference, _
                            "Endnote " & en.Index & " has identical content to endnote " & enFirstIdx, _
                            "Remove duplicate or differentiate content", _
                            en.Reference.Start, en.Reference.End, "possible_error"
                    Else
                        enDupDict.Add enCleanText, en.Index
                    End If
                End If
            End If

NextEnCombined:
        Next ei
    End If

    ' -- Cleanup Harts module-level caches -------------------------
    If hartsCount > 0 Then
        On Error Resume Next
        Application.Run "Rules_FootnoteHarts.ClearFootnoteCaches"
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
    End If

    ' -- Merge into master collection ------------------------------
    Dim fni As Long
    For fni = 1 To fnIssues.Count
        allIssues.Add fnIssues(fni)
    Next fni

    PerfTimerEnd "footnote_rules_combined"
    TraceStep "RunFootnoteRules", "completed: " & fnIssues.Count & " issues"
End Sub

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

    ' -- Initialise grouped report / comment suppression state --
    InitGroupedReportState

    ' -- Reset cancellation flag before any long work begins --
    ResetCancelRun

    ' -- Capture and suppress screen redraws for performance ----
    Dim wasScreenUpdating As Boolean
    wasScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error GoTo RunnerCleanup

    ' -- Build paragraph position cache (one scan, enables O(log N) lookups) --
    BuildParagraphCache doc

    ' -- Assess document complexity (once per run) --
    ComputeDocumentComplexity doc

    ' -- Precompute page-range character boundaries (one-time, cheap thereafter) --
    InitPageFilter doc

    ' -- Run fast prechecks (bracket balance, spelling presence, etc.) --
    Prechecks.RunPrechecks doc, config

    ' -- Whitelist rule first (populates whitelistDict) --
    If IsRuleEnabled(config, "custom_term_whitelist") Then
        PerfTimerStart "custom_term_whitelist"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Terms.Check_CustomTermWhitelist", doc)
        PerfTimerEnd "custom_term_whitelist"
    End If
    CheckCancellation

    ' -- Spellchecker (spelling + licence/license + check/cheque) --
    If IsRuleEnabled(config, "spellchecker") And Not Prechecks.SkipSpelling Then
        PerfTimerStart "spellchecker"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_Spelling", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_LicenceLicense", doc)
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Spelling.Check_CheckCheque", doc)
        PerfTimerEnd "spellchecker"
    End If

    CheckCancellation
    ' -- Consolidated paragraph-level rules (single pass) ----------
    ' Handles: repeated_words, spell_out_under_ten, double_spaces,
    '          triplicate_punctuation, dash_usage, bracket_integrity,
    '          double_commas, missing_space_after_dot,
    '          always_capitalise_terms, non_english_terms (anglicised + foreign)
    RunParagraphRules doc, config, allIssues

    CheckCancellation
    ' -- Number format rules (Range.Find based, not paragraph-level) --
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

    CheckCancellation
    ' -- Non-paragraph punctuation rules (Range.Find based) --------
    If IsRuleEnabled(config, "punctuation") Then
        PerfTimerStart "punctuation_find"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Punctuation.Check_SlashStyle", doc)
        PerfTimerEnd "punctuation_find"
    End If

    CheckCancellation
    ' -- Footnote rules (consolidated single pass) -----------------
    ' Handles: footnote integrity (sequence, placement, empty),
    '          footnotes-not-endnotes, Hart's rules (terminal stop,
    '          initial capital, abbreviation dictionary), and
    '          duplicate footnote detection.
    If IsRuleEnabled(config, "footnote_rules") Or _
       IsRuleEnabled(config, "duplicate_footnotes") Then
        RunFootnoteRules doc, config, allIssues
    End If

    CheckCancellation
    ' -- Brand names --
    If IsRuleEnabled(config, "brand_name_enforcement") Then
        PerfTimerStart "brand_name_enforcement"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_Brands.Check_BrandNameEnforcement", doc)
        PerfTimerEnd "brand_name_enforcement"
    End If

    CheckCancellation
    ' -- Mandated legal term forms (Range.Find based, not paragraph-level) --
    If IsRuleEnabled(config, "mandated_legal_term_forms") Then
        PerfTimerStart "mandated_legal_term_forms"
        AddIssuesToCollection allIssues, _
            TryRunRule("Rules_LegalTerms.Check_MandatedLegalTermForms", doc)
        PerfTimerEnd "mandated_legal_term_forms"
    End If

RunnerCleanup:
    ' -- 1. Capture the error that brought us here ---------------
    Dim wasCancelled As Boolean
    Dim hadUnexpectedErr As Boolean
    Dim savedErrNum As Long
    Dim savedErrDesc As String
    Dim savedErrSrc As String

    savedErrNum = Err.Number
    savedErrDesc = Err.Description
    savedErrSrc = Err.Source
    wasCancelled = (savedErrNum = ERR_RUN_CANCELLED)
    hadUnexpectedErr = (savedErrNum <> 0 And Not wasCancelled)
    If savedErrNum <> 0 Then Err.Clear

    ' Log unexpected errors so they are visible in the error log
    If hadUnexpectedErr Then
        ruleErrorCount = ruleErrorCount + 1
        ruleErrorLog = ruleErrorLog & "RunnerCleanup (Err " & savedErrNum & _
                       ": " & savedErrDesc & " [" & savedErrSrc & "])" & vbCrLf
    End If

    ' -- 2. Tear down caches (always) ----------------------------
    On Error Resume Next
    Prechecks.ClearPrechecks
    On Error GoTo 0

    ' -- 3. Restore application state (always) -------------------
    On Error Resume Next
    Application.ScreenUpdating = wasScreenUpdating
    Application.StatusBar = False   ' restore default status bar
    On Error GoTo 0

    ' -- 4. Post-processing: skip if cancelled or errored --------
    If Not wasCancelled And Not hadUnexpectedErr Then
        ' Filter out issues inside block quotes / quoted text
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

        ' Build grouped report data (must happen after all filtering)
        BuildSpellingGroups allIssues
        BuildFootnoteGroups allIssues

        ' Print performance summary
        If ENABLE_PROFILING Then
            Dim perfSummary As String
            perfSummary = GetPerformanceSummary()
            Debug.Print perfSummary
        End If

        ' Print debug summary
        PrintDebugSummary allIssues

        TraceStep "RunAllPleadingsRules", "total issues: " & allIssues.Count & _
                  ", rule errors: " & ruleErrorCount
        TraceExit "RunAllPleadingsRules", allIssues.Count & " issues"
    ElseIf wasCancelled Then
        TraceExit "RunAllPleadingsRules", "CANCELLED"
    Else
        TraceExit "RunAllPleadingsRules", "ERROR " & savedErrNum
    End If

    Set RunAllPleadingsRules = allIssues

    ' -- 5. Re-raise the original error AFTER all state is restored --
    If wasCancelled Then
        Err.Raise ERR_RUN_CANCELLED, "PleadingsEngine", "Run cancelled"
    ElseIf hadUnexpectedErr Then
        Err.Raise savedErrNum, savedErrSrc, savedErrDesc
    End If
End Function

' ============================================================
'  FILTER: Remove issues inside block quotes, cover pages,
'  and contents/table-of-contents pages.
'
'  Uses pre-built region arrays from BuildParagraphCache.
' ============================================================
Private Function FilterBlockQuoteIssues(doc As Document, _
                                         issues As Collection) As Collection
    TraceEnter "FilterBlockQuoteIssues"
    TraceStep "FilterBlockQuoteIssues", "input: " & issues.Count & " issues"
    Dim filtered As New Collection
    Dim i As Long

    ' -- Early exit if no regions detected -------------------------
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
        Dim inBQ As Boolean
        inBQ = False
        Dim j As Long
        For j = 0 To bqCount - 1
            If rs >= bqStarts(j) And rs < bqEnds(j) Then
                inBQ = True
                Exit For
            End If
        Next j
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
        If i Mod 20 = 0 Then CheckCancellation
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
                        gCommentsCreated = gCommentsCreated + 1
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
    Dim hlCancelled As Boolean
    Dim hlHadErr As Boolean
    Dim hlErrNum As Long
    Dim hlErrDesc As String
    Dim hlErrSrc As String

    hlErrNum = Err.Number
    hlErrDesc = Err.Description
    hlErrSrc = Err.Source
    hlCancelled = (hlErrNum = ERR_RUN_CANCELLED)
    hlHadErr = (hlErrNum <> 0 And Not hlCancelled)
    If hlErrNum <> 0 Then Err.Clear

    On Error Resume Next
    Application.ScreenUpdating = wasScreenUpdating
    Application.StatusBar = False
    On Error GoTo 0
    TraceExit "ApplyHighlights"
    If hlCancelled Then
        Err.Raise ERR_RUN_CANCELLED, "PleadingsEngine", "Run cancelled"
    ElseIf hlHadErr Then
        Err.Raise hlErrNum, hlErrSrc, hlErrDesc
    End If
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
        If i Mod 20 = 0 Then CheckCancellation
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
                    Dim thisRuleName As String
                    thisRuleName = CStr(GetIssueProp(finding, "RuleName"))

                    origStart = rng.Start
                    origLen = rng.End - rng.Start

                    ' --- FINDING OUTPUT MODE GATE (Section A) ---
                    ' Only OUTPUT_TRACKED_SAFE findings proceed to tracked edit
                    Dim outputMode As String
                    outputMode = GetFindingOutputMode(finding, doc)
                    If outputMode <> OUTPUT_TRACKED_SAFE Then
                        cntSkippedUnsafe = cntSkippedUnsafe + 1
                        gEditsSkippedUnsafe = gEditsSkippedUnsafe + 1
                        TraceStep "ApplyTrackedChanges", "SKIPPED output_mode=" & outputMode & _
                                  " i=" & i & " rule=" & thisRuleName
                        If addComments And outputMode = OUTPUT_COMMENT_ONLY And _
                           ShouldCreateCommentForRule(thisRuleName, finding) Then
                            TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                                "ApplyTrackedChanges", "mode-downgrade-comment i=" & i
                            gCommentsCreated = gCommentsCreated + 1
                        End If
                        GoTo NextApplyIssue
                    End If

                    ' --- UNSAFE AUTOFIX CATEGORY GATE (legacy, kept as belt-and-braces) ---
                    If IsUnsafeAutofixRule(thisRuleName) Then
                        cntSkippedUnsafe = cntSkippedUnsafe + 1
                        gEditsSkippedUnsafe = gEditsSkippedUnsafe + 1
                        TraceStep "ApplyTrackedChanges", "SKIPPED unsafe-category i=" & i & _
                                  " rule=" & thisRuleName
                        If addComments And ShouldCreateCommentForRule(thisRuleName, finding) Then
                            TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                                "ApplyTrackedChanges", "unsafe-category-comment i=" & i
                            gCommentsCreated = gCommentsCreated + 1
                        End If
                        GoTo NextApplyIssue
                    End If

                    ' Use ReplacementText only.  Suggestion is human-readable
                    ' prose and must NEVER be applied as literal replacement text.
                    sugText = ""
                    If Not HasReplacementText(finding) Then
                        TraceStep "ApplyTrackedChanges", "NO ReplacementText for i=" & i & _
                                  " rule=" & thisRuleName & "; comment-only"
                        If addComments And ShouldCreateCommentForRule(thisRuleName, finding) Then
                            TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                                "ApplyTrackedChanges", "no-replacement-comment i=" & i
                            gCommentsCreated = gCommentsCreated + 1
                        End If
                        GoTo NextApplyIssue
                    End If
                    sugText = CStr(GetIssueProp(finding, "ReplacementText"))

                    ' --- READ CURRENT TEXT AND VALIDATE ---
                    origText = ""
                    origText = rng.Text
                    If Err.Number <> 0 Then origText = "": Err.Clear

                    skipAmendment = False

                    ' --- STRONG ANCHOR GATE ---
                    ' Require: MatchedText non-empty, contains alphanumeric,
                    ' current rng.Text exactly equals MatchedText
                    Dim storedMatch As String
                    storedMatch = CStr(GetIssueProp(finding, "MatchedText"))

                    If Len(storedMatch) = 0 Then
                        skipAmendment = True
                        Debug.Print "STRONG_ANCHOR: Skipped -- empty MatchedText for rule=" & thisRuleName
                    ElseIf Not IsStrongTrackedAnchor(storedMatch) Then
                        skipAmendment = True
                        Debug.Print "STRONG_ANCHOR: Skipped -- no alphanumeric in '" & Left$(storedMatch, 30) & "'"
                    ElseIf Len(origText) > 0 And origText <> storedMatch Then
                        skipAmendment = True
                        Debug.Print "STRONG_ANCHOR: Skipped stale -- stored='" & Left$(storedMatch, 30) & "' actual='" & Left$(origText, 30) & "'"
                    End If

                    ' --- WHITESPACE-ONLY GATE ---
                    ' Never auto-apply findings whose MatchedText is only whitespace/punctuation
                    If Not skipAmendment And Len(origText) > 0 Then
                        If Not IsStrongTrackedAnchor(origText) Then
                            skipAmendment = True
                            Debug.Print "WHITESPACE GATE: Skipped -- origText has no alphanumeric"
                        End If
                    End If

                    ' For deletions (empty suggestion = delete the range)
                    If Not skipAmendment Then
                        If Len(sugText) = 0 And Len(origText) > 0 Then
                            For chIdx = 1 To Len(origText)
                                ch = Mid$(origText, chIdx, 1)
                                If (ch >= "A" And ch <= "Z") Or _
                                   (ch >= "a" And ch <= "z") Or _
                                   (ch >= "0" And ch <= "9") Or _
                                   ch = "." Then
                                    skipAmendment = True
                                    Debug.Print "WHITESPACE VALIDATION: Skipped deletion of '" & origText & "'"
                                    Exit For
                                End If
                            Next chIdx
                        End If
                    End If

                    ' For replacements, verify we are only changing whitespace
                    If Not skipAmendment Then
                        If Len(sugText) > 0 And Len(origText) > 0 Then
                            isOnlyWhitespace = True
                            For chIdx = 1 To Len(origText)
                                ch = Mid$(origText, chIdx, 1)
                                If ch <> " " And ch <> vbTab And ch <> ChrW(160) Then
                                    isOnlyWhitespace = False
                                    Exit For
                                End If
                            Next chIdx

                            ' Whitespace-only origText: never tracked-change
                            If isOnlyWhitespace Then
                                skipAmendment = True
                                Debug.Print "WHITESPACE GATE: Skipped whitespace-only origText"
                            End If

                            If Not skipAmendment And Not isOnlyWhitespace Then
                                If Len(sugText) < Len(origText) Then
                                    origHasPeriod = (InStr(1, origText, ".") > 0)
                                    sugHasPeriod = (InStr(1, sugText, ".") > 0)
                                    If origHasPeriod And Not sugHasPeriod Then
                                        skipAmendment = True
                                        Debug.Print "WHITESPACE VALIDATION: Skipped -- would remove period"
                                    End If
                                End If
                            End If
                        End If
                    End If

                    If skipAmendment Then
                        cntSkippedUnsafe = cntSkippedUnsafe + 1
                        gEditsSkippedUnsafe = gEditsSkippedUnsafe + 1
                        TraceStep "ApplyTrackedChanges", "SKIPPED amendment i=" & i & _
                                  " orig=""" & Left$(origText, 30) & """ sug=""" & Left$(sugText, 30) & """"
                        If addComments And ShouldCreateCommentForRule(thisRuleName, finding) Then
                            TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                                "ApplyTrackedChanges", "skip-comment i=" & i
                            gCommentsCreated = gCommentsCreated + 1
                        End If
                        GoTo NextApplyIssue
                    End If

                    ' --- UNICODE SAFETY: reject replacement char U+FFFD ---
                    If Not IsReplacementSafe(sugText) Then
                        cntSkippedUnsafe = cntSkippedUnsafe + 1
                        gEditsSkippedUnsafe = gEditsSkippedUnsafe + 1
                        TraceStep "ApplyTrackedChanges", "SKIPPED UNSAFE REPLACEMENT (U+FFFD) i=" & i
                        If addComments And ShouldCreateCommentForRule(thisRuleName, finding) Then
                            TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                                "ApplyTrackedChanges", "unsafe-replacement-comment i=" & i
                            gCommentsCreated = gCommentsCreated + 1
                        End If
                        GoTo NextApplyIssue
                    End If

                    ' Apply tracked change
                    TraceStep "ApplyTrackedChanges", "APPLYING i=" & i & _
                              " range=" & origStart & "-" & (origStart + origLen) & _
                              " orig=""" & Left$(origText, 30) & """ -> """ & Left$(sugText, 30) & """"
                    TrySetRangeText rng, sugText, _
                        "ApplyTrackedChanges", "apply i=" & i
                    cntApplied = cntApplied + 1
                    gTrackedEditsApplied = gTrackedEditsApplied + 1
                Else
                    cntCommentOnly = cntCommentOnly + 1
                    If addComments And ShouldCreateCommentForRule( _
                            CStr(GetIssueProp(finding, "RuleName")), finding) Then
                        TryAddComment doc, rng, BuildCommentText(finding), cmtRef, _
                            "ApplyTrackedChanges", "comment-only i=" & i
                        gCommentsCreated = gCommentsCreated + 1
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
    Dim tcCancelled As Boolean
    Dim tcHadErr As Boolean
    Dim tcErrNum As Long
    Dim tcErrDesc As String
    Dim tcErrSrc As String

    tcErrNum = Err.Number
    tcErrDesc = Err.Description
    tcErrSrc = Err.Source
    tcCancelled = (tcErrNum = ERR_RUN_CANCELLED)
    tcHadErr = (tcErrNum <> 0 And Not tcCancelled)
    If tcErrNum <> 0 Then Err.Clear

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
    If tcCancelled Then
        Err.Raise ERR_RUN_CANCELLED, "PleadingsEngine", "Run cancelled"
    ElseIf tcHadErr Then
        Err.Raise tcErrNum, tcErrSrc, tcErrDesc
    End If
End Sub

' ============================================================
'  PRIVATE: Build comment text from an issue dictionary
' ============================================================
Private Function BuildCommentText(ByVal finding As Object) As String
    Dim txt As String
    txt = GetIssueProp(finding, "Issue")

    Dim rn As String
    rn = LCase$(GetIssueProp(finding, "RuleName"))

    ' Repeated word: clean to "Repeated word 'X'" only
    If rn = "repeated_words" Then
        ' Strip " -- review context" and any trailing suggestion
        Dim dashPos As Long
        dashPos = InStr(1, txt, " -- ")
        If dashPos > 0 Then txt = Left$(txt, dashPos - 1)
        BuildCommentText = txt
        Exit Function
    End If

    ' Spacing rules: first sentence only, no suggestion tail
    If rn = "space_before_punct" Or rn = "double_spaces" Or _
       rn = "missing_space_after_dot" Then
        ' Truncate after first sentence-ending period
        Dim dotEnd As Long
        dotEnd = InStr(1, txt, ".")
        If dotEnd > 0 And dotEnd < Len(txt) Then
            txt = Left$(txt, dotEnd)
        End If
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

    ' Resolve document name and full path from explicit doc parameter
    Dim docName As String
    Dim docFullPath As String
    On Error Resume Next
    If Not doc Is Nothing Then
        docName = doc.Name
        If Len(doc.Path) > 0 Then
            docFullPath = doc.FullName
        Else
            docFullPath = doc.Name
        End If
    Else
        docName = "(no document)"
        docFullPath = "(no document)"
    End If
    If Err.Number <> 0 Then docName = "(unknown)": docFullPath = "(unknown)": Err.Clear
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

    ' Build page-range label for report
    Dim prStr As String
    prStr = pageRangeString
    If Len(prStr) = 0 Then prStr = "all"

    Print #fileNum, "{"
    Print #fileNum, "  ""document"": """ & EscJSON(docName) & ""","
    Print #fileNum, "  ""document_path"": """ & EscJSON(docFullPath) & ""","
    Print #fileNum, "  ""page_range"": """ & EscJSON(prStr) & ""","
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
    Print #fileNum, "    ""invalid_anchor_count"": " & invalidAnchorCount & ","
    Print #fileNum, "    ""grouped_spelling_count"": " & gGroupedSpellingCount & ","
    Print #fileNum, "    ""grouped_footnote_count"": " & gGroupedFootnoteCount & ","
    Print #fileNum, "    ""comments_created"": " & gCommentsCreated & ","
    Print #fileNum, "    ""tracked_edits_applied"": " & gTrackedEditsApplied & ","
    Print #fileNum, "    ""edits_skipped_unsafe"": " & gEditsSkippedUnsafe
    Print #fileNum, "  },"

    ' -- Grouped spelling pairs --
    Print #fileNum, "  ""grouped_spelling"": ["
    If Not gSpellingGroups Is Nothing Then
        If gSpellingGroups.Count > 0 Then
            Dim spGKeys As Variant
            spGKeys = gSpellingGroups.keys
            Dim spIdx As Long
            For spIdx = 0 To gSpellingGroups.Count - 1
                Dim spParts() As String
                spParts = Split(CStr(spGKeys(spIdx)), "|")
                Dim spFrom As String, spTo As String
                spFrom = spParts(0)
                If UBound(spParts) >= 1 Then spTo = spParts(1) Else spTo = ""
                Dim spExLoc As String
                spExLoc = ""
                If gSpellingExamples.Exists(CStr(spGKeys(spIdx))) Then
                    spExLoc = CStr(gSpellingExamples(spGKeys(spIdx)))
                End If
                Print #fileNum, "    {""from"": """ & EscJSON(spFrom) & """, ""to"": """ & _
                                EscJSON(spTo) & """, ""count"": " & gSpellingGroups(spGKeys(spIdx)) & _
                                ", ""examples"": """ & EscJSON(spExLoc) & """}" & _
                                IIf(spIdx < gSpellingGroups.Count - 1, ",", "")
            Next spIdx
        End If
    End If
    Print #fileNum, "  ]"

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

    ' Grouped spelling summary
    If gGroupedSpellingCount > SPELLING_COMMENT_THRESHOLD Then
        result = result & vbCrLf & "Spelling: " & gGroupedSpellingCount & _
                 " findings (" & gSpellingGroups.Count & " unique pairs) -- grouped in report" & vbCrLf
    End If

    ' Grouped footnote summary
    If gGroupedFootnoteCount > FOOTNOTE_COMMENT_THRESHOLD Then
        result = result & "Footnotes: " & gGroupedFootnoteCount & _
                 " findings -- grouped in report" & vbCrLf
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
    d.Add "duplicate_footnotes", "Duplicate Footnotes"
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

    ' Punctuation sub-rules -> "Punctuation Checker"
    Select Case rn
        Case "slash_style", "bracket_integrity", "hyphens", "dash_usage", _
             "double_commas", "space_before_punct", "missing_space_after_dot", _
             "triplicate_punctuation", "punctuation"
            GetUILabel = "Punctuation Checker"
            Exit Function
        Case "spellchecker", "spelling", "licence_license", "check_cheque"
            GetUILabel = "Spellchecker"
            Exit Function
        ' Footnote sub-rules -> "Footnote Rules"
        Case "footnote_integrity", "footnote_harts", _
             "footnote_terminal_full_stop", "footnote_initial_capital", _
             "footnote_abbreviation", "footnote_abbreviation_dictionary", _
             "footnotes_not_endnotes", "footnote_rules", "duplicate_footnotes"
            GetUILabel = "Footnote Rules"
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
'  be created for this rule. Much stricter than before:
'  - Spacing and dash rules: always suppressed
'  - Spelling/footnote: suppressed above threshold
'  - Duplicate comment text: suppressed above per-text limit
'  - Total comment cap: suppressed when exceeded
'  - Complex documents: bias further towards report-only
' ============================================================
Public Function ShouldCreateCommentForRule(ByVal ruleName As String, _
                                           Optional ByVal finding As Object = Nothing) As Boolean
    Dim rn As String
    rn = LCase$(ruleName)
    Dim bucket As String
    bucket = GetRuleBucket(rn)

    ' Gate 1: Spacing rules -- never create comments
    If bucket = "spacing" Then
        ShouldCreateCommentForRule = False
        Exit Function
    End If

    ' Gate 2: Dash/hyphen rules -- never create comments
    If bucket = "dash" Then
        ShouldCreateCommentForRule = False
        Exit Function
    End If

    ' Gate 2b: Footnote rules -- never create inline comments.
    ' Footnote findings are structural and overload documents.
    If bucket = "footnote" Then
        ShouldCreateCommentForRule = False
        Exit Function
    End If

    ' Gate 3: Trailing spaces -- never
    If rn = "trailing_spaces" Or rn = "trailing_space" Then
        ShouldCreateCommentForRule = False
        Exit Function
    End If

    ' Gate 4: Total comment cap exceeded
    If gCommentsCreated >= MAX_INLINE_TOTAL_COMMENTS Then
        ShouldCreateCommentForRule = False
        Exit Function
    End If

    ' Gate 5: Spelling -- per-bucket threshold
    If bucket = "spelling" Then
        If gGroupedSpellingCount > MAX_INLINE_SPELLING_COMMENTS Then
            ShouldCreateCommentForRule = False
            Exit Function
        End If
    End If

    ' Gate 6: (removed -- unreachable because Gate 2b already
    ' returns False for all footnote-bucket rules.)

    ' Gate 7: Check issue text for spacing sub-types emitted under other rule names
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

        ' Gate 8: Duplicate comment text suppression
        If Len(issText) > 0 And Not gCommentTextCounts Is Nothing Then
            Dim textKey As String
            textKey = Left$(issText, 120)
            If gCommentTextCounts.Exists(textKey) Then
                Dim dupCount As Long
                dupCount = CLng(gCommentTextCounts(textKey))
                If dupCount >= MAX_DUPLICATE_COMMENT_TEXT_PER_RUN Then
                    ShouldCreateCommentForRule = False
                    Exit Function
                End If
                gCommentTextCounts(textKey) = dupCount + 1
            Else
                gCommentTextCounts(textKey) = 1
            End If
        End If
    End If

    ' Gate 9: Complex document bias -- suppress less-important rule comments
    If gDocIsComplex Then
        ' In complex docs, only allow comments from high-value rules
        If Not IsCommentSafeRule(rn) Then
            ShouldCreateCommentForRule = False
            Exit Function
        End If
    End If

    ShouldCreateCommentForRule = True
End Function

' ============================================================
'  CONFIG DRIFT VALIDATION (development helper -- not user-facing)
'  Call from Immediate window: PleadingsEngine.ValidateConfigDrift
'  Compares config keys (InitRuleConfig) against display-name
'  keys (GetRuleDisplayNames) and prints mismatches.  Useful
'  after adding or renaming a rule to verify everything is wired.
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
        pageRangeString = ""
        Exit Sub
    End If
    pageRangeString = CStr(startPage) & "-" & CStr(endPage)
    Set pageRangeSet = CreateObject("Scripting.Dictionary")
    Dim pg As Long
    For pg = startPage To endPage
        pageRangeSet(pg) = True
    Next pg
End Sub

' ============================================================
'  PAGE-RANGE PARSING (Word-safe)
'  Supports:  5 | 3-7 | 3:7 | 1,3,5 | 1,3-5,8 | 5, 7-8, 9:30
'  Normalises en dash, em dash, minus sign to hyphen.
'  Colons also treated as range separators.
'  Parsed once per run -- not inside rule loops.
' ============================================================
Public Sub SetPageRangeFromString(ByVal spec As String)
    spec = Trim(spec)
    pageRangeString = spec  ' Preserve original for report metadata
    If Len(spec) = 0 Then
        Set pageRangeSet = Nothing
        Exit Sub
    End If

    ' Normalise input
    spec = NormalizePageRangeInput(spec)
    If Len(spec) = 0 Then
        Set pageRangeSet = Nothing
        Exit Sub
    End If

    ' Parse into page-number array
    Dim pages() As Long
    pages = ParsePageList(spec)

    ' Build dictionary from array
    Set pageRangeSet = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(pages) To UBound(pages)
        If pages(i) > 0 Then
            pageRangeSet(pages(i)) = True
        End If
    Next i

    ' If nothing valid was parsed, clear the set
    If pageRangeSet.Count = 0 Then
        Set pageRangeSet = Nothing
    End If
End Sub

' ============================================================
'  Normalise a page-range string for safe VBA parsing.
'  Strips control characters, collapses whitespace, normalises
'  en dash / em dash / minus sign to ASCII hyphen.
' ============================================================
Public Function NormalizePageRangeInput(ByVal s As String) As String
    s = CStr(s)
    s = Replace$(s, vbCr, "")
    s = Replace$(s, vbLf, "")
    s = Replace$(s, vbTab, "")
    s = Replace$(s, Chr$(160), " ")         ' non-breaking space
    s = Replace$(s, ChrW$(8211), "-")       ' en dash
    s = Replace$(s, ChrW$(8212), "-")       ' em dash
    s = Replace$(s, ChrW$(8722), "-")       ' minus sign
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop
    Do While InStr(s, ",,") > 0
        s = Replace$(s, ",,", ",")
    Loop
    NormalizePageRangeInput = s
End Function

' ============================================================
'  Parse a normalised page-range string into an array of page
'  numbers.  Supports comma-separated tokens where each token
'  is either a single number or a range separated by hyphen or
'  colon (inclusive).  Returns a 0-based Long array; element 0
'  is 0 when input is empty/invalid.
' ============================================================
Public Function ParsePageList(ByVal inputText As String) As Long()
    Dim parts() As String
    Dim subParts() As String
    Dim token As String
    Dim tmpList As Collection
    Dim arr() As Long
    Dim i As Long
    Dim p As Long
    Dim startPage As Long
    Dim endPage As Long
    Dim sepPos As Long
    Set tmpList = New Collection

    inputText = NormalizePageRangeInput(inputText)
    If Len(inputText) = 0 Then GoTo EmptyExit

    ' Reject inputs whose first non-space character is not a digit.
    ' Valid page ranges always start with a number; this catches
    ' leaked placeholder text like "e.g. 1,3,5-8,9:30".
    Dim firstNonSpace As Long
    Dim fc As String
    For firstNonSpace = 1 To Len(inputText)
        fc = Mid$(inputText, firstNonSpace, 1)
        If fc <> " " Then Exit For
    Next firstNonSpace
    If Not (fc >= "0" And fc <= "9") Then GoTo EmptyExit

    parts = Split(inputText, ",")
    For i = LBound(parts) To UBound(parts)
        token = Trim$(parts(i))
        If Len(token) = 0 Then GoTo NextPageToken

        ' Support both hyphen and colon as range separators
        sepPos = InStr(1, token, "-", vbBinaryCompare)
        If sepPos = 0 Then sepPos = InStr(1, token, ":", vbBinaryCompare)

        If sepPos > 0 Then
            ' Reject double separators inside a single token
            If InStr(sepPos + 1, token, "-", vbBinaryCompare) > 0 _
               Or InStr(sepPos + 1, token, ":", vbBinaryCompare) > 0 Then
                GoTo NextPageToken
            End If

            If Mid$(token, sepPos, 1) = "-" Then
                subParts = Split(token, "-")
            Else
                subParts = Split(token, ":")
            End If

            If UBound(subParts) = 1 Then
                If IsNumeric(Trim$(subParts(0))) And IsNumeric(Trim$(subParts(1))) Then
                    startPage = CLng(Trim$(subParts(0)))
                    endPage = CLng(Trim$(subParts(1)))
                    If startPage > 0 And endPage > 0 Then
                        If endPage < startPage Then
                            p = startPage
                            startPage = endPage
                            endPage = p
                        End If
                        For p = startPage To endPage
                            tmpList.Add p
                        Next p
                    End If
                End If
            End If
        Else
            ' Single page number
            If IsNumeric(token) Then
                p = CLng(token)
                If p > 0 Then tmpList.Add p
            End If
        End If
NextPageToken:
    Next i

    If tmpList.Count = 0 Then GoTo EmptyExit

    ReDim arr(0 To tmpList.Count - 1)
    For i = 1 To tmpList.Count
        arr(i - 1) = CLng(tmpList(i))
    Next i
    ParsePageList = arr
    Exit Function

EmptyExit:
    ReDim arr(0 To 0)
    arr(0) = 0
    ParsePageList = arr
End Function

' ============================================================
'  Check whether a page number is in a parsed page-number array.
'  Used by modules that need per-item page checks.
' ============================================================
Public Function IsPageSelected(ByVal pageNum As Long, ByRef selectedPages() As Long) As Boolean
    Dim i As Long
    If pageNum <= 0 Then Exit Function
    On Error GoTo SafeExit
    For i = LBound(selectedPages) To UBound(selectedPages)
        If selectedPages(i) = pageNum Then
            IsPageSelected = True
            Exit Function
        End If
    Next i
SafeExit:
End Function

Public Function GetPageRangeString() As String
    GetPageRangeString = pageRangeString
End Function

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
Public Function IsRuleEnabled(config As Object, _
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
    Dim afs As Boolean
    afs = False
    On Error Resume Next
    afs = CBool(GetIssueProp(finding, "AutoFixSafe"))
    If Err.Number <> 0 Then afs = False: Err.Clear
    On Error GoTo 0
    s = s & "      ""auto_fix_safe"": " & IIf(afs, "true", "false") & "," & vbCrLf
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
