Attribute VB_Name = "Rule09_DateTimeFormat"
' ============================================================
' Rule09_DateTimeFormat.bas
' Detects date and time format inconsistencies across the
' document. Identifies UK, US, and numeric date formats,
' determines the dominant style, and flags deviations.
' Also checks for mixed 12-hour / 24-hour time formats.
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "date_time_format"

' ── Helper: validate a month name ───────────────────────────
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

' ── Helper: search and collect date/time occurrences ────────
Private Sub FindWithWildcard(doc As Document, ByVal pattern As String, _
                              results As Collection, ByVal formatType As String)
    Dim rng As Range
    Set rng = doc.Content.Duplicate

    With rng.Find
        .ClearFormatting
        .Text = pattern
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .MatchCase = False
    End With

    Do While rng.Find.Execute
        If PleadingsEngine.IsInPageRange(rng) Then
            Dim info(0 To 3) As Variant
            info(0) = formatType
            info(1) = rng.Text
            info(2) = rng.Start
            info(3) = rng.End
            results.Add info
        End If
        rng.Collapse wdCollapseEnd
    Loop
End Sub

' ════════════════════════════════════════════════════════════
'  MAIN RULE FUNCTION
' ════════════════════════════════════════════════════════════
Public Function Check_DateTimeFormat(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' ══════════════════════════════════════════════════════════
    '  PASS 1: Find all date occurrences
    ' ══════════════════════════════════════════════════════════
    Dim dateFinds As New Collection
    Dim dateCounts As New Scripting.Dictionary
    dateCounts.Add "UK", 0
    dateCounts.Add "US", 0
    dateCounts.Add "numeric", 0

    ' ── UK format: "1 January 2024" or "12 March 2025" ──────
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

    ' ── US format: "January 1, 2024" or "March 12, 2025" ───
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

    ' ── Numeric format: "01/02/2024" or "1/2/24" ───────────
    Dim numResults As New Collection
    FindWithWildcard doc, "[0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4}", numResults, "numeric"

    For i = 1 To numResults.Count
        dateFinds.Add numResults(i)
        dateCounts("numeric") = dateCounts("numeric") + 1
    Next i

    ' ── Determine dominant date format ──────────────────────
    Dim dominantDate As String
    Dim maxDateCount As Long
    dominantDate = ""
    maxDateCount = 0
    Dim dk As Variant
    For Each dk In dateCounts.keys
        If dateCounts(dk) > maxDateCount Then
            maxDateCount = dateCounts(dk)
            dominantDate = CStr(dk)
        End If
    Next dk

    ' ── Flag non-dominant date formats ──────────────────────
    If maxDateCount > 0 Then
        Dim totalDateFormats As Long
        totalDateFormats = 0
        For Each dk In dateCounts.keys
            If dateCounts(dk) > 0 Then totalDateFormats = totalDateFormats + 1
        Next dk

        ' Only flag if there are mixed formats
        If totalDateFormats > 1 Then
            For i = 1 To dateFinds.Count
                Dim dInfo As Variant
                dInfo = dateFinds(i)
                Dim dType As String
                dType = CStr(dInfo(0))

                If dType <> dominantDate Then
                    Dim issueD As New PleadingsIssue
                    Dim rngD As Range
                    Set rngD = doc.Range(CLng(dInfo(2)), CLng(dInfo(3)))
                    Dim locD As String
                    locD = PleadingsEngine.GetLocationString(rngD, doc)

                    Dim suggestion As String
                    Select Case dominantDate
                        Case "UK"
                            suggestion = "Reformat to UK style (e.g., '1 January 2024')"
                        Case "US"
                            suggestion = "Reformat to US style (e.g., 'January 1, 2024')"
                        Case "numeric"
                            suggestion = "Reformat to numeric style (e.g., '01/01/2024')"
                    End Select

                    issueD.Init RULE_NAME, locD, _
                        "Inconsistent date format: '" & CStr(dInfo(1)) & _
                        "' uses " & dType & " format but dominant is " & dominantDate, _
                        suggestion, CLng(dInfo(2)), CLng(dInfo(3)), "error"
                    issues.Add issueD
                End If
            Next i
        End If
    End If

    ' ══════════════════════════════════════════════════════════
    '  PASS 2: Find time format inconsistencies
    ' ══════════════════════════════════════════════════════════
    Dim timeFinds As New Collection
    Dim timeCounts As New Scripting.Dictionary
    timeCounts.Add "12hr", 0
    timeCounts.Add "24hr", 0

    ' ── 12-hour format: "2:30 PM", "11:00 am" ──────────────
    Dim time12Results As New Collection
    FindWithWildcard doc, "[0-9]{1,2}:[0-9]{2} [AaPp][Mm]", time12Results, "12hr"

    For i = 1 To time12Results.Count
        timeFinds.Add time12Results(i)
        timeCounts("12hr") = timeCounts("12hr") + 1
    Next i

    ' ── 24-hour format: "14:30", "23:00" (hour >= 13) ──────
    Dim time24Results As New Collection
    FindWithWildcard doc, "[0-9]{2}:[0-9]{2}", time24Results, "24hr"

    For i = 1 To time24Results.Count
        Dim t24Info As Variant
        t24Info = time24Results(i)
        Dim t24Text As String
        t24Text = CStr(t24Info(1))

        ' Extract hour portion and check if >= 13
        Dim colonPos As Long
        colonPos = InStr(1, t24Text, ":")
        If colonPos > 0 Then
            Dim hourStr As String
            hourStr = Left$(t24Text, colonPos - 1)
            Dim hourVal As Long
            hourVal = 0
            If IsNumeric(hourStr) Then hourVal = CLng(hourStr)
            If hourVal >= 13 And hourVal <= 23 Then
                timeFinds.Add t24Info
                timeCounts("24hr") = timeCounts("24hr") + 1
            End If
        End If
    Next i

    ' ── Determine dominant time format and flag deviations ──
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

                If tType <> dominantTime Then
                    Dim issueT As New PleadingsIssue
                    Dim rngT As Range
                    Set rngT = doc.Range(CLng(tInfo(2)), CLng(tInfo(3)))
                    Dim locT As String
                    locT = PleadingsEngine.GetLocationString(rngT, doc)

                    Dim timeSugg As String
                    If dominantTime = "12hr" Then
                        timeSugg = "Use 12-hour format (e.g., '2:30 PM') for consistency"
                    Else
                        timeSugg = "Use 24-hour format (e.g., '14:30') for consistency"
                    End If

                    issueT.Init RULE_NAME, locT, _
                        "Inconsistent time format: '" & CStr(tInfo(1)) & _
                        "' uses " & tType & " format but dominant is " & dominantTime, _
                        timeSugg, CLng(tInfo(2)), CLng(tInfo(3)), "error"
                    issues.Add issueT
                End If
            Next i
        End If
    End If

    On Error GoTo 0
    Set Check_DateTimeFormat = issues
End Function
