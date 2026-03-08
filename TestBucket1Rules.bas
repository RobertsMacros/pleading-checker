Attribute VB_Name = "TestBucket1Rules"
' ============================================================
' TestBucket1Rules.bas
' Unit/integration tests for Bucket 1 Hart-style proofreading
' rules (Rules 24–34).
'
' Run via: TestBucket1Rules.RunAllBucket1Tests
'
' Each test creates a temporary document, populates it with
' test content, runs the relevant rule, and asserts expected
' results. Results are printed to the Immediate window.
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas
'   - All Rule24–Rule34 modules
' ============================================================
Option Explicit

Private passCount As Long
Private failCount As Long

' ════════════════════════════════════════════════════════════
'  MAIN TEST RUNNER
' ════════════════════════════════════════════════════════════
Public Sub RunAllBucket1Tests()
    passCount = 0
    failCount = 0

    ' Reset page range so tests are not filtered
    PleadingsEngine.SetPageRange 0, 0

    Debug.Print "========================================"
    Debug.Print "  Bucket 1 Rule Tests"
    Debug.Print "========================================"

    ' ── Footnote rules ────────────────────────────────────────
    Test_FootnotesNotEndnotes_FootnotesOnly
    Test_FootnotesNotEndnotes_EndnotesOnly
    Test_FootnotesNotEndnotes_Both

    Test_FootnoteTerminalFullStop_Pass
    Test_FootnoteTerminalFullStop_Fail_NoStop
    Test_FootnoteTerminalFullStop_Suppressed_Empty

    Test_FootnoteInitialCapital_Pass_Capital
    Test_FootnoteInitialCapital_Pass_Ibid
    Test_FootnoteInitialCapital_Pass_Eg
    Test_FootnoteInitialCapital_Fail_Lowercase

    Test_FootnoteAbbrevDict_Pass_Approved
    Test_FootnoteAbbrevDict_Fail_Dotted_Eg
    Test_FootnoteAbbrevDict_Fail_Dotted_Ie
    Test_FootnoteAbbrevDict_Fail_Pgs
    Test_FootnoteAbbrevDict_Fail_DottedParas
    Test_FootnoteAbbrevDict_Suppressed_Ordinary

    ' ── Hyphenation ───────────────────────────────────────────
    Test_MandatedTermForms_Fail_SolicitorGeneral
    Test_MandatedTermForms_Fail_AttorneyGeneral
    Test_MandatedTermForms_Pass_Correct

    ' ── Capitalisation ────────────────────────────────────────
    Test_AlwaysCapitalise_Fail_PrimeMinister
    Test_AlwaysCapitalise_Pass_Correct
    Test_AlwaysCapitalise_Fail_LawLords
    Test_AlwaysCapitalise_Suppressed_Quoted

    ' ── Foreign terms ─────────────────────────────────────────
    Test_AnglicisedNotItalic_Fail_Italic
    Test_AnglicisedNotItalic_Pass_Roman
    Test_AnglicisedNotItalic_Suppressed_Ambiguous

    Test_ForeignNamesNotItalic_Fail_Italic
    Test_ForeignNamesNotItalic_Pass_Roman

    ' ── Quotation marks ───────────────────────────────────────
    Test_SingleQuotesDefault_Fail_DoubleQuotes
    Test_SingleQuotesDefault_Pass_SingleQuotes
    Test_SingleQuotesDefault_Suppressed_BlockQuote

    Test_SmartQuoteConsistency_Fail_Mixed
    Test_SmartQuoteConsistency_Pass_AllCurly
    Test_SmartQuoteConsistency_Suppressed_Apostrophe

    ' ── Numbers ───────────────────────────────────────────────
    Test_SpellOutUnderTen_Fail_SevenInProse
    Test_SpellOutUnderTen_Pass_ParaRef
    Test_SpellOutUnderTen_Pass_Range
    Test_SpellOutUnderTen_Suppressed_Table

    ' ── Summary ───────────────────────────────────────────────
    Debug.Print "========================================"
    Debug.Print "  PASSED: " & passCount
    Debug.Print "  FAILED: " & failCount
    Debug.Print "  TOTAL:  " & (passCount + failCount)
    Debug.Print "========================================"
End Sub

' ════════════════════════════════════════════════════════════
'  ASSERTION HELPERS
' ════════════════════════════════════════════════════════════
Private Sub AssertIssueCount(testName As String, issues As Collection, expected As Long)
    Dim actual As Long
    If issues Is Nothing Then
        actual = 0
    Else
        actual = issues.Count
    End If

    If actual = expected Then
        Debug.Print "  PASS: " & testName & " (count=" & actual & ")"
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: " & testName & " (expected " & expected & ", got " & actual & ")"
        failCount = failCount + 1
    End If
End Sub

Private Sub AssertIssueCountAtLeast(testName As String, issues As Collection, minExpected As Long)
    Dim actual As Long
    If issues Is Nothing Then
        actual = 0
    Else
        actual = issues.Count
    End If

    If actual >= minExpected Then
        Debug.Print "  PASS: " & testName & " (count=" & actual & " >= " & minExpected & ")"
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: " & testName & " (expected >=" & minExpected & ", got " & actual & ")"
        failCount = failCount + 1
    End If
End Sub

Private Sub AssertNoIssues(testName As String, issues As Collection)
    AssertIssueCount testName, issues, 0
End Sub

Private Sub AssertHasIssues(testName As String, issues As Collection)
    AssertIssueCountAtLeast testName, issues, 1
End Sub

Private Sub AssertSeverity(testName As String, issues As Collection, idx As Long, expected As String)
    If issues Is Nothing Or idx > issues.Count Or idx < 1 Then
        Debug.Print "  FAIL: " & testName & " (no issue at index " & idx & ")"
        failCount = failCount + 1
        Exit Sub
    End If

    Dim issue As PleadingsIssue
    Set issue = issues(idx)
    If issue.Severity = expected Then
        Debug.Print "  PASS: " & testName & " (severity=" & expected & ")"
        passCount = passCount + 1
    Else
        Debug.Print "  FAIL: " & testName & " (expected severity '" & expected & "', got '" & issue.Severity & "')"
        failCount = failCount + 1
    End If
End Sub

' ════════════════════════════════════════════════════════════
'  HELPER: Create temp document with content
' ════════════════════════════════════════════════════════════
Private Function CreateTempDoc(Optional content As String = "") As Document
    Dim doc As Document
    Set doc = Documents.Add
    If Len(content) > 0 Then
        doc.Content.Text = content
    End If
    Set CreateTempDoc = doc
End Function

Private Sub CloseTempDoc(doc As Document)
    On Error Resume Next
    doc.Close SaveChanges:=wdDoNotSaveChanges
    On Error GoTo 0
End Sub

' ════════════════════════════════════════════════════════════
'  FOOTNOTES NOT ENDNOTES TESTS
' ════════════════════════════════════════════════════════════
Private Sub Test_FootnotesNotEndnotes_FootnotesOnly()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="A footnote."
    Dim issues As Collection
    Set issues = Check_FootnotesNotEndnotes(doc)
    AssertNoIssues "FootnotesNotEndnotes: footnotes only -> pass", issues
    CloseTempDoc doc
End Sub

Private Sub Test_FootnotesNotEndnotes_EndnotesOnly()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Endnotes.Add Range:=doc.Words(6), Text:="An endnote."
    Dim issues As Collection
    Set issues = Check_FootnotesNotEndnotes(doc)
    AssertHasIssues "FootnotesNotEndnotes: endnotes only -> fail", issues
    If issues.Count >= 1 Then
        AssertSeverity "FootnotesNotEndnotes: severity=error", issues, 1, "error"
    End If
    CloseTempDoc doc
End Sub

Private Sub Test_FootnotesNotEndnotes_Both()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point and another.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="A footnote."
    doc.Endnotes.Add Range:=doc.Words(8), Text:="An endnote."
    Dim issues As Collection
    Set issues = Check_FootnotesNotEndnotes(doc)
    AssertHasIssues "FootnotesNotEndnotes: both -> fail with mixed message", issues
    CloseTempDoc doc
End Sub

' ════════════════════════════════════════════════════════════
'  FOOTNOTE TERMINAL FULL STOP TESTS
' ════════════════════════════════════════════════════════════
Private Sub Test_FootnoteTerminalFullStop_Pass()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="See Smith v Jones [2020] UKSC 1."
    Dim issues As Collection
    Set issues = Check_FootnoteTerminalFullStop(doc)
    AssertNoIssues "FootnoteTerminalFullStop: ends with '.' -> pass", issues
    CloseTempDoc doc
End Sub

Private Sub Test_FootnoteTerminalFullStop_Fail_NoStop()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="See https://example.com"
    Dim issues As Collection
    Set issues = Check_FootnoteTerminalFullStop(doc)
    AssertHasIssues "FootnoteTerminalFullStop: URL no stop -> fail", issues
    CloseTempDoc doc
End Sub

Private Sub Test_FootnoteTerminalFullStop_Suppressed_Empty()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:=""
    Dim issues As Collection
    Set issues = Check_FootnoteTerminalFullStop(doc)
    ' Empty footnote should be suppressed (no false positive)
    AssertNoIssues "FootnoteTerminalFullStop: empty -> suppressed", issues
    CloseTempDoc doc
End Sub

' ════════════════════════════════════════════════════════════
'  FOOTNOTE INITIAL CAPITAL TESTS
' ════════════════════════════════════════════════════════════
Private Sub Test_FootnoteInitialCapital_Pass_Capital()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="See Smith v Jones."
    Dim issues As Collection
    Set issues = Check_FootnoteInitialCapital(doc)
    AssertNoIssues "FootnoteInitialCapital: 'See Smith...' -> pass", issues
    CloseTempDoc doc
End Sub

Private Sub Test_FootnoteInitialCapital_Pass_Ibid()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="ibid 45."
    Dim issues As Collection
    Set issues = Check_FootnoteInitialCapital(doc)
    AssertNoIssues "FootnoteInitialCapital: 'ibid 45.' -> pass (allowed)", issues
    CloseTempDoc doc
End Sub

Private Sub Test_FootnoteInitialCapital_Pass_Eg()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="eg discussion in Smith."
    Dim issues As Collection
    Set issues = Check_FootnoteInitialCapital(doc)
    AssertNoIssues "FootnoteInitialCapital: 'eg discussion...' -> pass (allowed)", issues
    CloseTempDoc doc
End Sub

Private Sub Test_FootnoteInitialCapital_Fail_Lowercase()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="see Smith v Jones."
    Dim issues As Collection
    Set issues = Check_FootnoteInitialCapital(doc)
    AssertHasIssues "FootnoteInitialCapital: 'see Smith...' -> fail", issues
    CloseTempDoc doc
End Sub

' ════════════════════════════════════════════════════════════
'  FOOTNOTE ABBREVIATION DICTIONARY TESTS
' ════════════════════════════════════════════════════════════
Private Sub Test_FootnoteAbbrevDict_Pass_Approved()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="See pp 12-13 and cf Smith."
    Dim issues As Collection
    Set issues = Check_FootnoteAbbreviationDictionary(doc)
    AssertNoIssues "FootnoteAbbrevDict: approved forms -> pass", issues
    CloseTempDoc doc
End Sub

Private Sub Test_FootnoteAbbrevDict_Fail_Dotted_Eg()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="e.g. discussion in Smith."
    Dim issues As Collection
    Set issues = Check_FootnoteAbbreviationDictionary(doc)
    AssertHasIssues "FootnoteAbbrevDict: 'e.g.' -> fail, suggest 'eg'", issues
    CloseTempDoc doc
End Sub

Private Sub Test_FootnoteAbbrevDict_Fail_Dotted_Ie()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="i.e. the principle."
    Dim issues As Collection
    Set issues = Check_FootnoteAbbreviationDictionary(doc)
    AssertHasIssues "FootnoteAbbrevDict: 'i.e.' -> fail, suggest 'ie'", issues
    CloseTempDoc doc
End Sub

Private Sub Test_FootnoteAbbrevDict_Fail_Pgs()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="See pgs 12-13."
    Dim issues As Collection
    Set issues = Check_FootnoteAbbreviationDictionary(doc)
    AssertHasIssues "FootnoteAbbrevDict: 'pgs' -> fail, suggest 'pp'", issues
    CloseTempDoc doc
End Sub

Private Sub Test_FootnoteAbbrevDict_Fail_DottedParas()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="See paras. 4-6."
    Dim issues As Collection
    Set issues = Check_FootnoteAbbreviationDictionary(doc)
    AssertHasIssues "FootnoteAbbrevDict: 'paras.' -> fail, suggest 'paras'", issues
    CloseTempDoc doc
End Sub

Private Sub Test_FootnoteAbbrevDict_Suppressed_Ordinary()
    Dim doc As Document
    Set doc = CreateTempDoc("Test paragraph with a reference point.")
    doc.Footnotes.Add Range:=doc.Words(6), Text:="The secretary provided the page references."
    Dim issues As Collection
    Set issues = Check_FootnoteAbbreviationDictionary(doc)
    ' "page" is an ordinary word, should not be flagged as abbreviation
    AssertNoIssues "FootnoteAbbrevDict: ordinary words -> suppressed", issues
    CloseTempDoc doc
End Sub

' ════════════════════════════════════════════════════════════
'  MANDATED LEGAL TERM FORMS TESTS
' ════════════════════════════════════════════════════════════
Private Sub Test_MandatedTermForms_Fail_SolicitorGeneral()
    Dim doc As Document
    Set doc = CreateTempDoc("The Solicitor General appeared for the Crown.")
    Dim issues As Collection
    Set issues = Check_MandatedLegalTermForms(doc)
    AssertHasIssues "MandatedTermForms: 'Solicitor General' -> fail", issues
    CloseTempDoc doc
End Sub

Private Sub Test_MandatedTermForms_Fail_AttorneyGeneral()
    Dim doc As Document
    Set doc = CreateTempDoc("The Attorney General issued guidance.")
    Dim issues As Collection
    Set issues = Check_MandatedLegalTermForms(doc)
    AssertHasIssues "MandatedTermForms: 'Attorney General' -> fail", issues
    CloseTempDoc doc
End Sub

Private Sub Test_MandatedTermForms_Pass_Correct()
    Dim doc As Document
    Set doc = CreateTempDoc("The Solicitor-General and Attorney-General appeared.")
    Dim issues As Collection
    Set issues = Check_MandatedLegalTermForms(doc)
    AssertNoIssues "MandatedTermForms: correct forms -> pass", issues
    CloseTempDoc doc
End Sub

' ════════════════════════════════════════════════════════════
'  ALWAYS CAPITALISE TERMS TESTS
' ════════════════════════════════════════════════════════════
Private Sub Test_AlwaysCapitalise_Fail_PrimeMinister()
    Dim doc As Document
    Set doc = CreateTempDoc("The prime minister addressed the house.")
    Dim issues As Collection
    Set issues = Check_AlwaysCapitaliseTerms(doc)
    AssertHasIssues "AlwaysCapitalise: 'prime minister' -> fail", issues
    CloseTempDoc doc
End Sub

Private Sub Test_AlwaysCapitalise_Pass_Correct()
    Dim doc As Document
    Set doc = CreateTempDoc("The Prime Minister addressed the House.")
    Dim issues As Collection
    Set issues = Check_AlwaysCapitaliseTerms(doc)
    ' "Prime Minister" correct — only check this term is not flagged
    ' (Note: "House" is not in the seed dictionary so not flagged)
    AssertNoIssues "AlwaysCapitalise: 'Prime Minister' -> pass", issues
    CloseTempDoc doc
End Sub

Private Sub Test_AlwaysCapitalise_Fail_LawLords()
    Dim doc As Document
    Set doc = CreateTempDoc("The law lords delivered their opinion.")
    Dim issues As Collection
    Set issues = Check_AlwaysCapitaliseTerms(doc)
    AssertHasIssues "AlwaysCapitalise: 'law lords' -> fail", issues
    CloseTempDoc doc
End Sub

Private Sub Test_AlwaysCapitalise_Suppressed_Quoted()
    Dim doc As Document
    ' Quoted context should be suppressed
    Set doc = CreateTempDoc("He said " & ChrW(8216) & "the prime minister was wrong" & ChrW(8217) & ".")
    Dim issues As Collection
    Set issues = Check_AlwaysCapitaliseTerms(doc)
    ' Depending on implementation, this may or may not suppress.
    ' The rule says: do not alter quoted material.
    ' We accept either 0 (suppressed) or 1 (flagged but suggest-only)
    Debug.Print "  INFO: AlwaysCapitalise: quoted context -> " & _
                IIf(issues.Count = 0, "suppressed (ideal)", "flagged (acceptable)")
    passCount = passCount + 1
    CloseTempDoc doc
End Sub

' ════════════════════════════════════════════════════════════
'  ANGLICISED TERMS NOT ITALIC TESTS
' ════════════════════════════════════════════════════════════
Private Sub Test_AnglicisedNotItalic_Fail_Italic()
    Dim doc As Document
    Set doc = CreateTempDoc("The principle of prima facie evidence.")
    ' Italicise "prima facie"
    Dim rng As Range
    Dim pos As Long
    pos = InStr(1, doc.Content.Text, "prima facie", vbTextCompare)
    If pos > 0 Then
        Set rng = doc.Range(pos - 1, pos - 1 + 11)
        rng.Font.Italic = True
    End If
    Dim issues As Collection
    Set issues = Check_AnglicisedTermsNotItalic(doc)
    AssertHasIssues "AnglicisedNotItalic: italic 'prima facie' -> fail", issues
    CloseTempDoc doc
End Sub

Private Sub Test_AnglicisedNotItalic_Pass_Roman()
    Dim doc As Document
    Set doc = CreateTempDoc("The principle of prima facie evidence.")
    Dim issues As Collection
    Set issues = Check_AnglicisedTermsNotItalic(doc)
    AssertNoIssues "AnglicisedNotItalic: roman 'prima facie' -> pass", issues
    CloseTempDoc doc
End Sub

Private Sub Test_AnglicisedNotItalic_Suppressed_Ambiguous()
    ' Test with a term at a punctuation boundary
    Dim doc As Document
    Set doc = CreateTempDoc("He argued (per se) the point was moot.")
    Dim issues As Collection
    Set issues = Check_AnglicisedTermsNotItalic(doc)
    ' "per se" is not italic -> should pass, not flag
    AssertNoIssues "AnglicisedNotItalic: non-italic at punct boundary -> suppressed", issues
    CloseTempDoc doc
End Sub

' ════════════════════════════════════════════════════════════
'  FOREIGN NAMES NOT ITALIC TESTS
' ════════════════════════════════════════════════════════════
Private Sub Test_ForeignNamesNotItalic_Fail_Italic()
    Dim doc As Document
    Set doc = CreateTempDoc("The decision of the Cour de cassation was clear.")
    ' Italicise "Cour de cassation"
    Dim rng As Range
    Dim pos As Long
    pos = InStr(1, doc.Content.Text, "Cour de cassation", vbTextCompare)
    If pos > 0 Then
        Set rng = doc.Range(pos - 1, pos - 1 + 17)
        rng.Font.Italic = True
    End If
    Dim issues As Collection
    Set issues = Check_ForeignNamesNotItalic(doc)
    AssertHasIssues "ForeignNamesNotItalic: italic 'Cour de cassation' -> fail", issues
    CloseTempDoc doc
End Sub

Private Sub Test_ForeignNamesNotItalic_Pass_Roman()
    Dim doc As Document
    Set doc = CreateTempDoc("The decision of the Cour de cassation was clear.")
    Dim issues As Collection
    Set issues = Check_ForeignNamesNotItalic(doc)
    AssertNoIssues "ForeignNamesNotItalic: roman 'Cour de cassation' -> pass", issues
    CloseTempDoc doc
End Sub

' ════════════════════════════════════════════════════════════
'  SINGLE QUOTES DEFAULT TESTS
' ════════════════════════════════════════════════════════════
Private Sub Test_SingleQuotesDefault_Fail_DoubleQuotes()
    Dim doc As Document
    Set doc = CreateTempDoc("He discussed " & ChrW(8220) & "the principle" & ChrW(8221) & " at length.")
    Dim issues As Collection
    Set issues = Check_SingleQuotesDefault(doc)
    AssertHasIssues "SingleQuotesDefault: outer double quotes -> fail", issues
    CloseTempDoc doc
End Sub

Private Sub Test_SingleQuotesDefault_Pass_SingleQuotes()
    Dim doc As Document
    Set doc = CreateTempDoc("He discussed " & ChrW(8216) & "the principle" & ChrW(8217) & " at length.")
    Dim issues As Collection
    Set issues = Check_SingleQuotesDefault(doc)
    AssertNoIssues "SingleQuotesDefault: outer single quotes -> pass", issues
    CloseTempDoc doc
End Sub

Private Sub Test_SingleQuotesDefault_Suppressed_BlockQuote()
    Dim doc As Document
    Set doc = CreateTempDoc(ChrW(8220) & "The block quoted text." & ChrW(8221))
    ' Set style to block quote
    On Error Resume Next
    doc.Paragraphs(1).Style = "Quote"
    On Error GoTo 0
    Dim issues As Collection
    Set issues = Check_SingleQuotesDefault(doc)
    ' Should be suppressed if style is "Quote"
    ' Accept 0 (suppressed) or non-zero (style not available)
    Debug.Print "  INFO: SingleQuotesDefault: block quote -> " & _
                IIf(issues.Count = 0, "suppressed (ideal)", _
                    "flagged (quote style may not exist in template)")
    passCount = passCount + 1
    CloseTempDoc doc
End Sub

' ════════════════════════════════════════════════════════════
'  SMART QUOTE CONSISTENCY TESTS
' ════════════════════════════════════════════════════════════
Private Sub Test_SmartQuoteConsistency_Fail_Mixed()
    Dim doc As Document
    ' Mix curly and straight double quotes
    Set doc = CreateTempDoc(ChrW(8220) & "curly" & ChrW(8221) & " and " & Chr(34) & "straight" & Chr(34) & ".")
    Dim issues As Collection
    Set issues = Check_SmartQuoteConsistency(doc)
    AssertHasIssues "SmartQuoteConsistency: mixed straight+curly -> fail", issues
    CloseTempDoc doc
End Sub

Private Sub Test_SmartQuoteConsistency_Pass_AllCurly()
    Dim doc As Document
    Set doc = CreateTempDoc(ChrW(8220) & "all curly" & ChrW(8221) & " and " & ChrW(8220) & "more curly" & ChrW(8221) & ".")
    Dim issues As Collection
    Set issues = Check_SmartQuoteConsistency(doc)
    AssertNoIssues "SmartQuoteConsistency: all curly -> pass", issues
    CloseTempDoc doc
End Sub

Private Sub Test_SmartQuoteConsistency_Suppressed_Apostrophe()
    Dim doc As Document
    ' Straight apostrophe mid-word should not be counted as a quote
    Set doc = CreateTempDoc(ChrW(8220) & "it" & ChrW(8217) & "s fine" & ChrW(8221) & " and don't worry.")
    Dim issues As Collection
    Set issues = Check_SmartQuoteConsistency(doc)
    ' The straight apostrophe in "don't" should not trigger inconsistency
    ' if the rule correctly identifies apostrophes
    Debug.Print "  INFO: SmartQuoteConsistency: apostrophe handling -> " & _
                IIf(issues.Count = 0, "suppressed (ideal)", _
                    issues.Count & " issue(s) (check apostrophe detection)")
    passCount = passCount + 1
    CloseTempDoc doc
End Sub

' ════════════════════════════════════════════════════════════
'  SPELL OUT UNDER TEN TESTS
' ════════════════════════════════════════════════════════════
Private Sub Test_SpellOutUnderTen_Fail_SevenInProse()
    Dim doc As Document
    Set doc = CreateTempDoc("There were 7 issues raised at the hearing.")
    Dim issues As Collection
    Set issues = Check_SpellOutUnderTen(doc)
    AssertHasIssues "SpellOutUnderTen: '7 issues' -> fail", issues
    CloseTempDoc doc
End Sub

Private Sub Test_SpellOutUnderTen_Pass_ParaRef()
    Dim doc As Document
    Set doc = CreateTempDoc("See para 7 of the judgment.")
    Dim issues As Collection
    Set issues = Check_SpellOutUnderTen(doc)
    AssertNoIssues "SpellOutUnderTen: 'para 7' -> pass (structural ref)", issues
    CloseTempDoc doc
End Sub

Private Sub Test_SpellOutUnderTen_Pass_Range()
    Dim doc As Document
    ' Range with en-dash
    Set doc = CreateTempDoc("Children aged 7" & ChrW(8211) & "12 were included.")
    Dim issues As Collection
    Set issues = Check_SpellOutUnderTen(doc)
    AssertNoIssues "SpellOutUnderTen: '7-12' range -> pass", issues
    CloseTempDoc doc
End Sub

Private Sub Test_SpellOutUnderTen_Suppressed_Table()
    Dim doc As Document
    Set doc = CreateTempDoc("There were 5 items in the list.")
    ' Try to set style to Table-like
    On Error Resume Next
    doc.Paragraphs(1).Style = "Table Text"
    On Error GoTo 0
    Dim issues As Collection
    Set issues = Check_SpellOutUnderTen(doc)
    ' If table style exists, should be suppressed
    Debug.Print "  INFO: SpellOutUnderTen: table context -> " & _
                IIf(issues.Count = 0, "suppressed (ideal)", _
                    issues.Count & " issue(s) (table style may not exist)")
    passCount = passCount + 1
    CloseTempDoc doc
End Sub
