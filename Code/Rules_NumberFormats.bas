Attribute VB_Name = "Rules_NumberFormats"
' ============================================================
' Rules_NumberFormats.bas
' Combined module for number/date/currency format rules:
'   - Rule09: Date and time format consistency
'   - Rule19: Currency and number format consistency
'
' Public functions:
'   Check_DateTimeFormat        (Rule09)
'   Check_CurrencyNumberFormat  (Rule19)
'
' Dependencies:
'   - TextAnchoring.bas (FindAll, AddIssue, SafeRange,
'     GetDateFormatPref, IsInPageRange)
' ============================================================
Option Explicit

' -- Rule name constants ---------------------------------------
Private Const RULE_NAME_DATE_TIME As String = "date_time_format"
Private Const RULE_NAME_CURRENCY As String = "currency_number_format"

' -- Currency format category constants (Rule19) ---------------
Private Const FMT_WORDS As String = "words"
Private Const FMT_ABBREVIATED As String = "abbreviated"
Private Const FMT_FULL_NUMERIC As String = "full_numeric"
Private Const FMT_ISO_PREFIX As String = "iso_prefix"


' ============================================================
'  PRIVATE HELPERS  -  Rule09 (Date/Time)
' ============================================================

' -- Helper: validate a month name -----------------------------
Private Function IsValidMonth(ByVal monthName As String) As Boolean
    Dim months As Variant
    months = Array("January", "February", "March", "April", "May", "June", _
                   "July", "August", "September", "October", "November", "December")
    Dim m As Variant
    For Each m In months
        If StrComp(monthName, CStr(m), vbTextCompare) = 0 Then
            IsValidMonth = True
            Exit Function
        End If
    Next m
    IsValidMonth = False
End Function

' -- Helper: search and collect date/time occurrences ----------
Private Sub FindWithWildcard(doc As Document, ByVal pattern As String, results As Collection, ByVal formatType As String)
    Dim matches As Collection
    Set matches = TextAnchoring.FindAll(doc, pattern, False, False, True)
    Dim i As Long
    For i = 1 To matches.Count
        Dim m As Variant: m = matches(i)
        Dim info(0 To 3) As Variant
        info(0) = formatType
        info(1) = CStr(m(2))
        info(2) = CLng(m(0))
        info(3) = CLng(m(1))
        results.Add info
    Next i
End Sub

' -- Helper: check if a time match looks like a clause reference,
'  ratio, date component, or other non-time pattern.
'  Examines characters before and after the HH:MM match.
'  NOTE: creates Range objects for single characters, but this runs
'  only on the small set of time-pattern matches, not per-paragraph.
' ----------------------------------------------------------------
Private Function LooksLikeNonTimeContext(doc As Document, _
        ByVal matchStart As Long, ByVal matchEnd As Long) As Boolean
    LooksLikeNonTimeContext = False
    On Error Resume Next

    ' Check character before the match
    If matchStart > 0 Then
        Dim bRng As Range
        Set bRng = doc.Range(matchStart - 1, matchStart)
        If Err.Number = 0 Then
            Dim bc As String
            bc = bRng.Text
            If Err.Number = 0 Then
                ' Preceded by letter -> probably part of a word or reference
                If (bc >= "A" And bc <= "Z") Or (bc >= "a" And bc <= "z") Then
                    LooksLikeNonTimeContext = True
                    Err.Clear: On Error GoTo 0: Exit Function
                End If
                ' Preceded by another digit -> could be ratio like 1:12:45
                If bc >= "0" And bc <= "9" Then
                    ' Check two chars back for another colon (chained ratio)
                    If matchStart > 1 Then
                        Dim b2Rng As Range
                        Set b2Rng = doc.Range(matchStart - 2, matchStart - 1)
                        If Err.Number = 0 Then
                            Dim b2c As String
                            b2c = b2Rng.Text
                            If b2c = ":" Or b2c = "." Then
                                LooksLikeNonTimeContext = True
                                Err.Clear: On Error GoTo 0: Exit Function
                            End If
                        Else
                            Err.Clear
                        End If
                    End If
                End If
            Else
                Err.Clear
            End If
        Else
            Err.Clear
        End If
    End If

    ' Check character after the match
    If matchEnd < doc.Content.End Then
        Dim aRng As Range
        Set aRng = doc.Range(matchEnd, matchEnd + 1)
        If Err.Number = 0 Then
            Dim ac As String
            ac = aRng.Text
            If Err.Number = 0 Then
                ' Followed by a colon or dot+digit -> ratio or version number
                If ac = ":" Then
                    LooksLikeNonTimeContext = True
                    Err.Clear: On Error GoTo 0: Exit Function
                End If
                If ac = "." Then
                    If matchEnd + 1 < doc.Content.End Then
                        Dim a2Rng As Range
                        Set a2Rng = doc.Range(matchEnd + 1, matchEnd + 2)
                        If Err.Number = 0 Then
                            Dim a2c As String
                            a2c = a2Rng.Text
                            If a2c >= "0" And a2c <= "9" Then
                                LooksLikeNonTimeContext = True
                                Err.Clear: On Error GoTo 0: Exit Function
                            End If
                        Else
                            Err.Clear
                        End If
                    End If
                End If
                ' Followed by a letter -> part of a word
                If (ac >= "A" And ac <= "Z") Or (ac >= "a" And ac <= "z") Then
                    LooksLikeNonTimeContext = True
                    Err.Clear: On Error GoTo 0: Exit Function
                End If
            Else
                Err.Clear
            End If
        Else
            Err.Clear
        End If
    End If

    On Error GoTo 0
End Function

' ============================================================
'  PRIVATE HELPERS  -  Rule19 (Currency/Number)
' ============================================================

' -- Check format consistency for a single symbol --------------
'  Searches for words, abbreviated, and full_numeric formats,
'  determines the dominant format, and flags minorities.
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

    ' -- Search for "words" format: symbol + digits + space + word --
    Dim wordMatches As Collection
    Set wordMatches = TextAnchoring.FindAll(doc, sym & "[0-9.]@" & " [a-z]@", False, False, True)
    Dim wi As Long
    For wi = 1 To wordMatches.Count
        Dim wm As Variant: wm = wordMatches(wi)
        If IsMagnitudeWord(LCase(CStr(wm(2)))) Then
            wordsCount = wordsCount + 1
            wordsRanges.Add TextAnchoring.SafeRange(doc, CLng(wm(0)), CLng(wm(1)))
        End If
    Next wi

    ' -- Search for "abbreviated" format: symbol + digits + m/bn/k --
    Dim abbrMatches As Collection
    Set abbrMatches = TextAnchoring.FindAll(doc, sym & "[0-9.]@[mbk]", False, False, True)
    Dim ai As Long
    For ai = 1 To abbrMatches.Count
        Dim am As Variant: am = abbrMatches(ai)
        abbrCount = abbrCount + 1
        abbrRanges.Add TextAnchoring.SafeRange(doc, CLng(am(0)), CLng(am(1)))
    Next ai

    ' -- Search for "full_numeric" format: symbol + digits with commas --
    Dim numMatches As Collection
    Set numMatches = TextAnchoring.FindAll(doc, sym & "[0-9,.]@", False, False, True)
    Dim ni As Long
    For ni = 1 To numMatches.Count
        Dim nm As Variant: nm = numMatches(ni)
        Dim numText As String: numText = CStr(nm(2))
        If InStr(numText, ",") > 0 And Len(numText) >= 5 Then
            numericCount = numericCount + 1
            numericRanges.Add TextAnchoring.SafeRange(doc, CLng(nm(0)), CLng(nm(1)))
        End If
    Next ni

    ' -- Determine dominant format and flag minorities ----------
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

' -- Check ISO code prefixed amounts ---------------------------
'  Searches for patterns like "GBP 1,500" or "USD 25.00"
Private Sub CheckISOCodeFormat(doc As Document, isoCode As String, ByRef issues As Collection)
    Dim matches As Collection
    Set matches = TextAnchoring.FindAll(doc, isoCode & " [0-9]@", False, True, True)
    Dim i As Long
    For i = 1 To matches.Count
        Dim m As Variant: m = matches(i)
        Dim rng As Range: Set rng = TextAnchoring.SafeRange(doc, CLng(m(0)), CLng(m(1)))
        TextAnchoring.AddIssue issues, RULE_NAME_CURRENCY, doc, rng, "ISO code format used: '" & CStr(m(2)) & "'", "Consider using symbol notation for consistency", CLng(m(0)), CLng(m(1)), "possible_error"
    Next i
End Sub

' -- Flag all ranges in a minority format collection -----------
Private Sub FlagMinorityRanges(doc As Document, ranges As Collection, symLabel As String, minorityFmt As String, dominantFmt As String, ByRef issues As Collection)
    Dim i As Long
    For i = 1 To ranges.Count
        Dim rng As Range: Set rng = ranges(i)
        If Not rng Is Nothing Then
            TextAnchoring.AddIssue issues, RULE_NAME_CURRENCY, doc, rng, symLabel & " amount uses '" & minorityFmt & "' format: '" & rng.Text & "'", "Use '" & dominantFmt & "' format for consistency (dominant style)", rng.Start, rng.End, "error"
        End If
    Next i
End Sub

' -- Check if matched text contains a magnitude word -----------
Private Function IsMagnitudeWord(ByVal txt As String) As Boolean
    Dim lTxt As String
    lTxt = LCase(txt)

    IsMagnitudeWord = (InStr(lTxt, "million") > 0) Or _
                      (InStr(lTxt, "billion") > 0) Or _
                      (InStr(lTxt, "thousand") > 0) Or _
                      (InStr(lTxt, "hundred") > 0) Or _
                      (InStr(lTxt, "trillion") > 0)
End Function

' ============================================================
'  PUBLIC FUNCTIONS
' ============================================================

' ================================================================
'  Rule09: Check_DateTimeFormat
'  Detects date and time format inconsistencies across the
'  document. Identifies UK, US, and numeric date formats,
'  determines the dominant style, and flags deviations.
'  Also checks for mixed 12-hour / 24-hour time formats.
'
'  24-hour detection recognises 00:00 through 23:59 with
'  context filtering to exclude clause references and ratios.
' ================================================================
Public Function Check_DateTimeFormat(doc As Document) As Collection
    Dim issues As New Collection

    ' ==========================================================
    '  PASS 1: Find all date occurrences
    ' ==========================================================
    Dim dateFinds As New Collection
    Dim dateCounts As Object
    Set dateCounts = CreateObject("Scripting.Dictionary")
    dateCounts.Add "UK", 0
    dateCounts.Add "US", 0
    dateCounts.Add "numeric", 0

    ' -- UK format: "1 January 2024" or "12 March 2025" ------
    ' VBA wildcard: one or two digits, space, word, space, four digits
    Dim ukResults As New Collection
    FindWithWildcard doc, "[0-9]{1,2} [A-Z][a-z]{2,} [0-9]{4}", ukResults, "UK"

    ' Validate UK results (check month name)
    Dim ukItem As Variant
    Dim i As Long
    For i = 1 To ukResults.Count
        Dim ukInfo As Variant
        ukInfo = ukResults(i)
        Dim ukText As String
        ukText = CStr(ukInfo(1))

        ' Extract month name (between first and last space)
        Dim parts() As String
        parts = Split(ukText, " ")
        If UBound(parts) >= 2 Then
            If IsValidMonth(parts(1)) Then
                dateFinds.Add ukInfo
                dateCounts("UK") = dateCounts("UK") + 1
            End If
        End If
    Next i

    ' -- US format: "January 1, 2024" or "March 12, 2025" ----
    Dim usResults As New Collection
    FindWithWildcard doc, "[A-Z][a-z]{2,} [0-9]{1,2}, [0-9]{4}", usResults, "US"

    For i = 1 To usResults.Count
        Dim usInfo As Variant
        usInfo = usResults(i)
        Dim usText As String
        usText = CStr(usInfo(1))

        parts = Split(usText, " ")
        If UBound(parts) >= 0 Then
            If IsValidMonth(parts(0)) Then
                dateFinds.Add usInfo
                dateCounts("US") = dateCounts("US") + 1
            End If
        End If
    Next i

    ' -- Numeric format: "01/02/2024" or "1/2/24" -------------
    Dim numResults As New Collection
    FindWithWildcard doc, "[0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4}", numResults, "numeric"

    For i = 1 To numResults.Count
        dateFinds.Add numResults(i)
        dateCounts("numeric") = dateCounts("numeric") + 1
    Next i

    ' -- Determine dominant date format ------------------------
    Dim dominantDate As String
    Dim maxDateCount As Long
    Dim dk As Variant

    ' Check user preference first
    Dim datePref As String
    datePref = TextAnchoring.GetDateFormatPref()

    If datePref = "UK" Or datePref = "US" Then
        ' User has set a preference -- use it as dominant
        dominantDate = datePref
        maxDateCount = dateCounts(datePref)
    Else
        ' AUTO mode: pick the most frequent format
        dominantDate = ""
        maxDateCount = 0
        For Each dk In dateCounts.keys
            If dateCounts(dk) > maxDateCount Then
                maxDateCount = dateCounts(dk)
                dominantDate = CStr(dk)
            End If
        Next dk
    End If

    ' -- Flag non-dominant date formats ------------------------
    If maxDateCount > 0 Then
        Dim totalDateFormats As Long
        totalDateFormats = 0
        For Each dk In dateCounts.keys
            If dateCounts(dk) > 0 Then totalDateFormats = totalDateFormats + 1
        Next dk

        ' Flag if there are mixed formats, or if a preference is set
        If totalDateFormats > 1 Or (datePref = "UK" Or datePref = "US") Then
            For i = 1 To dateFinds.Count
                Dim dInfo As Variant
                dInfo = dateFinds(i)
                Dim dType As String
                dType = CStr(dInfo(0))

                If dType <> dominantDate Then
                    Dim rngD As Range
                    Set rngD = TextAnchoring.SafeRange(doc, CLng(dInfo(2)), CLng(dInfo(3)))
                    If rngD Is Nothing Then GoTo NextDateFind
                    Dim suggestion As String
                    Select Case dominantDate
                        Case "UK": suggestion = "Reformat to UK style (e.g., '1 January 2024')"
                        Case "US": suggestion = "Reformat to US style (e.g., 'January 1, 2024')"
                        Case "numeric": suggestion = "Reformat to numeric style (e.g., '01/01/2024')"
                    End Select
                    TextAnchoring.AddIssue issues, RULE_NAME_DATE_TIME, doc, rngD, "Inconsistent date format: '" & CStr(dInfo(1)) & "' uses " & dType & " format but dominant is " & dominantDate, suggestion, CLng(dInfo(2)), CLng(dInfo(3)), "error"
                End If
NextDateFind:
            Next i
        End If
    End If

    ' ==========================================================
    '  PASS 2: Find time format inconsistencies
    '
    '  12-hour: explicit AM/PM marker (e.g. 2:30 PM, 11:00 am)
    '  24-hour: HH:MM where HH is 00-23, no AM/PM follows,
    '           and context does not suggest clause ref or ratio.
    ' ==========================================================
    Dim timeFinds As New Collection
    Dim timeCounts As Object
    Set timeCounts = CreateObject("Scripting.Dictionary")
    timeCounts.Add "12hr", 0
    timeCounts.Add "24hr", 0

    ' -- 12-hour format: "2:30 PM", "11:00 am" ----------------
    Dim time12Results As New Collection
    FindWithWildcard doc, "[0-9]{1,2}:[0-9]{2} [AaPp][Mm]", time12Results, "12hr"

    For i = 1 To time12Results.Count
        timeFinds.Add time12Results(i)
        timeCounts("12hr") = timeCounts("12hr") + 1
    Next i

    ' Also catch dot-separated 12hr times: "2.30 pm"
    Dim time12DotResults As New Collection
    FindWithWildcard doc, "[0-9]{1,2}.[0-9]{2} [AaPp][Mm]", time12DotResults, "12hr"

    For i = 1 To time12DotResults.Count
        timeFinds.Add time12DotResults(i)
        timeCounts("12hr") = timeCounts("12hr") + 1
    Next i

    ' -- 24-hour format: HH:MM (00:00 through 23:59) ----------
    '  Search for two-digit colon two-digit patterns.
    '  Filter: must be valid 00-23 hour and 00-59 minute.
    '  Exclude matches followed by AM/PM (those are 12-hour).
    '  Exclude matches in non-time context (clause refs, ratios).
    Dim time24Results As New Collection
    FindWithWildcard doc, "[0-9]{2}:[0-9]{2}", time24Results, "24hr"

    For i = 1 To time24Results.Count
        Dim t24Info As Variant
        t24Info = time24Results(i)
        Dim t24Text As String
        t24Text = CStr(t24Info(1))

        ' Parse hour and minute
        Dim colonPos As Long
        colonPos = InStr(1, t24Text, ":")
        If colonPos > 0 Then
            Dim hourStr As String
            hourStr = Left$(t24Text, colonPos - 1)
            Dim minStr As String
            minStr = Mid$(t24Text, colonPos + 1)
            Dim hourVal As Long
            Dim minVal As Long
            hourVal = -1
            minVal = -1
            If IsNumeric(hourStr) Then hourVal = CLng(hourStr)
            If IsNumeric(minStr) Then minVal = CLng(minStr)

            ' Valid time: hour 0-23, minute 0-59
            If hourVal >= 0 And hourVal <= 23 And minVal >= 0 And minVal <= 59 Then
                Dim is24hrTime As Boolean
                is24hrTime = True

                ' Check whether AM/PM follows (with or without space)
                ' to avoid double-counting 12-hour times
                Dim peekEnd As Long
                peekEnd = CLng(t24Info(3)) + 4
                On Error Resume Next
                If peekEnd > doc.Content.End Then peekEnd = doc.Content.End
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                If peekEnd > CLng(t24Info(3)) Then
                    Dim peekRng As Range
                    On Error Resume Next
                    Set peekRng = doc.Range(CLng(t24Info(3)), peekEnd)
                    Dim peekTxt As String
                    peekTxt = ""
                    peekTxt = UCase$(peekRng.Text)
                    If Err.Number <> 0 Then peekTxt = "": Err.Clear
                    On Error GoTo 0
                    ' Followed by AM/PM (with or without space) = 12-hour
                    If Len(peekTxt) >= 2 Then
                        If Left$(peekTxt, 2) = "AM" Or Left$(peekTxt, 2) = "PM" Then
                            is24hrTime = False
                        ElseIf Len(peekTxt) >= 3 Then
                            If Mid$(peekTxt, 2, 2) = "AM" Or Mid$(peekTxt, 2, 2) = "PM" Then
                                is24hrTime = False
                            End If
                        End If
                    End If
                End If

                ' Context check: exclude clause refs, ratios, etc.
                If is24hrTime Then
                    If LooksLikeNonTimeContext(doc, CLng(t24Info(2)), CLng(t24Info(3))) Then
                        is24hrTime = False
                    End If
                End If

                ' Classify: hours 13-23 or 00 are definite 24-hour.
                ' Hours 01-12 without AM/PM are ambiguous and should not
                ' drive the dominant-style count (but are still collected
                ' so they can be flagged if a clear dominant emerges).
                If is24hrTime Then
                    If hourVal >= 13 Or hourVal = 0 Then
                        ' Definite 24-hour: counts toward dominance
                        timeFinds.Add t24Info
                        timeCounts("24hr") = timeCounts("24hr") + 1
                    Else
                        ' Ambiguous (01:00-12:59 without AM/PM):
                        ' Collect for possible flagging but mark as "ambiguous"
                        ' so it does NOT influence the dominant format.
                        Dim ambigInfo(0 To 3) As Variant
                        ambigInfo(0) = "ambiguous"
                        ambigInfo(1) = t24Info(1)
                        ambigInfo(2) = t24Info(2)
                        ambigInfo(3) = t24Info(3)
                        timeFinds.Add ambigInfo
                        ' Do NOT increment timeCounts("24hr")
                    End If
                End If
            End If
        End If
    Next i

    ' -- Determine dominant time format and flag deviations ----
    Dim dominantTime As String
    Dim maxTimeCount As Long
    dominantTime = ""
    maxTimeCount = 0
    For Each dk In timeCounts.keys
        If timeCounts(dk) > maxTimeCount Then
            maxTimeCount = timeCounts(dk)
            dominantTime = CStr(dk)
        End If
    Next dk

    If maxTimeCount > 0 Then
        Dim totalTimeFormats As Long
        totalTimeFormats = 0
        For Each dk In timeCounts.keys
            If timeCounts(dk) > 0 Then totalTimeFormats = totalTimeFormats + 1
        Next dk

        If totalTimeFormats > 1 Then
            For i = 1 To timeFinds.Count
                Dim tInfo As Variant
                tInfo = timeFinds(i)
                Dim tType As String
                tType = CStr(tInfo(0))

                ' Skip ambiguous times: they don't conflict with anything
                If tType = "ambiguous" Then GoTo NextTimeFind
                If tType <> dominantTime Then
                    Dim rngT As Range
                    Set rngT = TextAnchoring.SafeRange(doc, CLng(tInfo(2)), CLng(tInfo(3)))
                    If rngT Is Nothing Then GoTo NextTimeFind
                    Dim timeSugg As String
                    If dominantTime = "12hr" Then timeSugg = "Use 12-hour format (e.g., '2:30 PM') for consistency" Else timeSugg = "Use 24-hour format (e.g., '14:30') for consistency"
                    TextAnchoring.AddIssue issues, RULE_NAME_DATE_TIME, doc, rngT, "Inconsistent time format: '" & CStr(tInfo(1)) & "' uses " & tType & " format but dominant is " & dominantTime, timeSugg, CLng(tInfo(2)), CLng(tInfo(3)), "error"
                End If
NextTimeFind:
            Next i
        End If
    End If

    Set Check_DateTimeFormat = issues
End Function


' ================================================================
'  Rule19: Check_CurrencyNumberFormat
'  Detects inconsistent currency/number formatting across
'  the document. Checks symbol-prefixed amounts (GBP, USD, EUR)
'  and ISO-code-prefixed amounts, then flags minority format
'  usage.
' ================================================================
Public Function Check_CurrencyNumberFormat(doc As Document) As Collection
    Dim issues As New Collection
    Dim symbols As Variant
    Dim symLabels As Variant
    Dim i As Long

    ' Primary currency symbols to check
    symbols = Array(ChrW(163), "$", ChrW(8364))   ' GBP, USD, EUR
    symLabels = Array("GBP", "USD", "EUR")

    ' -- Check each symbol for format consistency --------------
    For i = LBound(symbols) To UBound(symbols)
        CheckSymbolConsistency doc, CStr(symbols(i)), CStr(symLabels(i)), issues
    Next i

    ' -- Check ISO code prefixed amounts -----------------------
    Dim isoCodes As Variant
    isoCodes = Array("GBP", "USD", "EUR", "JPY", "AUD", "CAD", "CHF", _
                     "BTC", "ETH", "USDT", "USDC", "BNB", "XRP", "SOL", "ADA", "DOGE")

    For i = LBound(isoCodes) To UBound(isoCodes)
        CheckISOCodeFormat doc, CStr(isoCodes(i)), issues
    Next i

    Set Check_CurrencyNumberFormat = issues
End Function