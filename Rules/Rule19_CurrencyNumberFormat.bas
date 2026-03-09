Attribute VB_Name = "Rule19_CurrencyNumberFormat"
' ============================================================
' Rule19_CurrencyNumberFormat.bas
' Proofreading rule: detects inconsistent currency/number
' formatting across the document. Checks symbol-prefixed
' amounts (GBP, USD, EUR) and ISO-code-prefixed amounts,
' then flags minority format usage.
'
' Format categories:
'   words        - e.g. "$1.5 million"
'   abbreviated  - e.g. "$1.5m"
'   full_numeric - e.g. "$1,500,000"
'   iso_prefix   - e.g. "GBP 1,500"
'
' Dependencies:
'   - PleadingsIssue.cls
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "currency_number_format"

' ── Format category constants ───────────────────────────────
Private Const FMT_WORDS As String = "words"
Private Const FMT_ABBREVIATED As String = "abbreviated"
Private Const FMT_FULL_NUMERIC As String = "full_numeric"
Private Const FMT_ISO_PREFIX As String = "iso_prefix"

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Function Check_CurrencyNumberFormat(doc As Document) As Collection
    Dim issues As New Collection
    Dim symbols As Variant
    Dim symLabels As Variant
    Dim i As Long

    ' Primary currency symbols to check
    symbols = Array(ChrW(163), "$", ChrW(8364))   ' GBP, USD, EUR
    symLabels = Array("GBP", "USD", "EUR")

    ' ── Check each symbol for format consistency ────────────
    For i = LBound(symbols) To UBound(symbols)
        CheckSymbolConsistency doc, CStr(symbols(i)), CStr(symLabels(i)), issues
    Next i

    ' ── Check ISO code prefixed amounts ─────────────────────
    Dim isoCodes As Variant
    isoCodes = Array("GBP", "USD", "EUR", "JPY", "AUD", "CAD", "CHF", _
                     "BTC", "ETH", "USDT", "USDC", "BNB", "XRP", "SOL", "ADA", "DOGE")

    For i = LBound(isoCodes) To UBound(isoCodes)
        CheckISOCodeFormat doc, CStr(isoCodes(i)), issues
    Next i

    Set Check_CurrencyNumberFormat = issues
End Function

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check format consistency for a single symbol
'  Searches for words, abbreviated, and full_numeric formats,
'  determines the dominant format, and flags minorities.
' ════════════════════════════════════════════════════════════
Private Sub CheckSymbolConsistency(doc As Document, _
                                    sym As String, _
                                    symLabel As String, _
                                    ByRef issues As Collection)
    Dim wordsCount As Long
    Dim abbrCount As Long
    Dim numericCount As Long
    Dim wordsRanges As Collection
    Dim abbrRanges As Collection
    Dim numericRanges As Collection

    Set wordsRanges = New Collection
    Set abbrRanges = New Collection
    Set numericRanges = New Collection

    ' ── Search for "words" format: symbol + digits + space + word ──
    ' Pattern: e.g. £[0-9.]@ [a-z]@  (wildcard)
    Dim rng As Range
    Dim wordPattern As String
    wordPattern = sym & "[0-9.]@" & " [a-z]@"

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = wordPattern
        .MatchWildcards = True
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Do
        On Error Resume Next
        Dim found As Boolean
        found = rng.Find.Execute
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Do
        End If
        On Error GoTo 0

        If Not found Then Exit Do

        ' Validate that the trailing word is a magnitude word
        Dim matchText As String
        matchText = LCase(rng.Text)
        If IsMagnitudeWord(matchText) Then
            If PleadingsEngine.IsInPageRange(rng) Then
                wordsCount = wordsCount + 1
                wordsRanges.Add doc.Range(rng.Start, rng.End)
            End If
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    ' ── Search for "abbreviated" format: symbol + digits + m/bn/k ──
    Dim abbrPattern As String
    abbrPattern = sym & "[0-9.]@[mbk]"

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = abbrPattern
        .MatchWildcards = True
        .MatchCase = False
        .MatchWholeWord = False
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
            abbrCount = abbrCount + 1
            abbrRanges.Add doc.Range(rng.Start, rng.End)
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    ' ── Search for "full_numeric" format: symbol + digits with commas ──
    Dim numPattern As String
    numPattern = sym & "[0-9,.]@"

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = numPattern
        .MatchWildcards = True
        .MatchCase = False
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Do
        On Error Resume Next
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do

        ' Only count as full_numeric if it contains a comma and is long enough
        Dim numText As String
        numText = rng.Text
        If InStr(numText, ",") > 0 And Len(numText) >= 5 Then
            If PleadingsEngine.IsInPageRange(rng) Then
                numericCount = numericCount + 1
                numericRanges.Add doc.Range(rng.Start, rng.End)
            End If
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop

    ' ── Determine dominant format and flag minorities ──────
    Dim totalFormats As Long
    totalFormats = 0
    If wordsCount > 0 Then totalFormats = totalFormats + 1
    If abbrCount > 0 Then totalFormats = totalFormats + 1
    If numericCount > 0 Then totalFormats = totalFormats + 1

    ' Only flag if more than one format is in use
    If totalFormats < 2 Then Exit Sub

    ' Find the dominant format
    Dim domFormat As String
    Dim domCount As Long
    domFormat = FMT_WORDS: domCount = wordsCount
    If abbrCount > domCount Then domFormat = FMT_ABBREVIATED: domCount = abbrCount
    If numericCount > domCount Then domFormat = FMT_FULL_NUMERIC: domCount = numericCount

    ' Flag minority: words
    If wordsCount > 0 And domFormat <> FMT_WORDS Then
        FlagMinorityRanges doc, wordsRanges, symLabel, FMT_WORDS, domFormat, issues
    End If

    ' Flag minority: abbreviated
    If abbrCount > 0 And domFormat <> FMT_ABBREVIATED Then
        FlagMinorityRanges doc, abbrRanges, symLabel, FMT_ABBREVIATED, domFormat, issues
    End If

    ' Flag minority: full_numeric
    If numericCount > 0 And domFormat <> FMT_FULL_NUMERIC Then
        FlagMinorityRanges doc, numericRanges, symLabel, FMT_FULL_NUMERIC, domFormat, issues
    End If
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check ISO code prefixed amounts
'  Searches for patterns like "GBP 1,500" or "USD 25.00"
' ════════════════════════════════════════════════════════════
Private Sub CheckISOCodeFormat(doc As Document, _
                                isoCode As String, _
                                ByRef issues As Collection)
    Dim rng As Range
    Dim isoPattern As String
    Dim issue As PleadingsIssue
    Dim locStr As String

    ' Search for ISO code followed by space and number
    isoPattern = isoCode & " [0-9]@"

    Set rng = doc.Content.Duplicate
    With rng.Find
        .ClearFormatting
        .Text = isoPattern
        .MatchWildcards = True
        .MatchCase = True
        .MatchWholeWord = False
        .Wrap = wdFindStop
        .Forward = True
    End With

    Dim isoCount As Long
    isoCount = 0

    Do
        On Error Resume Next
        Dim found As Boolean
        found = rng.Find.Execute
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0

        If Not found Then Exit Do

        If PleadingsEngine.IsInPageRange(rng) Then
            isoCount = isoCount + 1

            ' Flag ISO prefix usage as informational (possible_error)
            ' since mixing ISO codes with symbol notation is inconsistent
            On Error Resume Next
            locStr = PleadingsEngine.GetLocationString(rng, doc)
            If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
            On Error GoTo 0

            Set issue = New PleadingsIssue
            issue.Init RULE_NAME, _
                       locStr, _
                       "ISO code format used: '" & rng.Text & "'", _
                       "Consider using symbol notation for consistency", _
                       rng.Start, _
                       rng.End, _
                       "possible_error"
            issues.Add issue
        End If

        On Error Resume Next
        rng.Collapse wdCollapseEnd
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Do
        On Error GoTo 0
    Loop
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Flag all ranges in a minority format collection
' ════════════════════════════════════════════════════════════
Private Sub FlagMinorityRanges(doc As Document, _
                                ranges As Collection, _
                                symLabel As String, _
                                minorityFmt As String, _
                                dominantFmt As String, _
                                ByRef issues As Collection)
    Dim i As Long
    Dim rng As Range
    Dim issue As PleadingsIssue
    Dim locStr As String

    For i = 1 To ranges.Count
        Set rng = ranges(i)

        On Error Resume Next
        locStr = PleadingsEngine.GetLocationString(rng, doc)
        If Err.Number <> 0 Then locStr = "unknown location": Err.Clear
        On Error GoTo 0

        Set issue = New PleadingsIssue
        issue.Init RULE_NAME, _
                   locStr, _
                   symLabel & " amount uses '" & minorityFmt & "' format: '" & rng.Text & "'", _
                   "Use '" & dominantFmt & "' format for consistency (dominant style)", _
                   rng.Start, _
                   rng.End, _
                   "error"
        issues.Add issue
    Next i
End Sub

' ════════════════════════════════════════════════════════════
'  PRIVATE: Check if matched text contains a magnitude word
' ════════════════════════════════════════════════════════════
Private Function IsMagnitudeWord(ByVal txt As String) As Boolean
    Dim lTxt As String
    lTxt = LCase(txt)

    IsMagnitudeWord = (InStr(lTxt, "million") > 0) Or _
                      (InStr(lTxt, "billion") > 0) Or _
                      (InStr(lTxt, "thousand") > 0) Or _
                      (InStr(lTxt, "hundred") > 0) Or _
                      (InStr(lTxt, "trillion") > 0)
End Function

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunCurrencyNumberFormat()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Currency Number Format"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_CurrencyNumberFormat(doc)

    ' Apply results with tracked changes (UK magic circle default)
    PleadingsEngine.ApplyIssuesToDocument doc, issues

    Application.ScreenUpdating = True

    MsgBox "Found " & issues.Count & " issue(s).", _
           vbInformation, "Currency Number Format"
End Sub
