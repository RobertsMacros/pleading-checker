Attribute VB_Name = "Rule10_InlineListFormat"
' ============================================================
' Rule10_InlineListFormat.bas
' Checks inline list formatting consistency: separator style
' (semicolon, comma, none), conjunction usage ("and"/"or"
' before final item), and ending punctuation. Handles (a),
' (i)/(ii)/(iii), and (1)/(2)/(3) marker styles.
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "inline_list_format"

' ── Marker pattern types ────────────────────────────────────
Private Const MARKER_LETTER As String = "letter"   ' (a), (b), (c)
Private Const MARKER_ROMAN As String = "roman"     ' (i), (ii), (iii)
Private Const MARKER_NUMBER As String = "number"   ' (1), (2), (3)

' ── Helper: detect marker type from content between parens ──
Private Function GetMarkerType(ByVal content As String) As String
    If Len(content) = 0 Then
        GetMarkerType = ""
        Exit Function
    End If

    ' Single lowercase letter: (a)-(z)
    If Len(content) = 1 And content Like "[a-z]" Then
        GetMarkerType = MARKER_LETTER
        Exit Function
    End If

    ' Numeric: (1), (2), (12)
    If IsNumeric(content) Then
        GetMarkerType = MARKER_NUMBER
        Exit Function
    End If

    ' Roman numeral: all chars are i, v, x, l, c, d, m
    Dim allRoman As Boolean
    Dim ci As Long
    allRoman = True
    For ci = 1 To Len(content)
        If Not (Mid$(content, ci, 1) Like "[ivxlcdm]") Then
            allRoman = False
            Exit For
        End If
    Next ci
    If allRoman Then
        GetMarkerType = MARKER_ROMAN
        Exit Function
    End If

    GetMarkerType = ""
End Function

' ── Helper: find all inline list markers in a paragraph ─────
' Returns Collection of Array(markerPos, markerText, markerContent, markerType)
Private Function FindMarkersInPara(ByVal paraText As String) As Collection
    Dim markers As New Collection
    Dim pos As Long
    Dim openParen As Long
    Dim closeParen As Long
    Dim content As String
    Dim mType As String

    pos = 1
    Do While pos <= Len(paraText)
        openParen = InStr(pos, paraText, "(")
        If openParen = 0 Then Exit Do

        closeParen = InStr(openParen + 1, paraText, ")")
        If closeParen = 0 Then Exit Do
        If closeParen - openParen > 6 Then
            ' Too long to be a list marker
            pos = openParen + 1
            GoTo ContinueSearch
        End If

        content = Mid$(paraText, openParen + 1, closeParen - openParen - 1)
        mType = GetMarkerType(content)

        If Len(mType) > 0 Then
            Dim info(0 To 3) As Variant
            info(0) = openParen         ' position in paragraph text
            info(1) = Mid$(paraText, openParen, closeParen - openParen + 1) ' full marker text
            info(2) = content           ' content between parens
            info(3) = mType             ' marker type
            markers.Add info
        End If

        pos = closeParen + 1
ContinueSearch:
    Loop

    Set FindMarkersInPara = markers
End Function

' ── Helper: detect separator before a marker ────────────────
' Looks at text between previous marker's end and current marker's start
Private Function DetectSeparator(ByVal textBetween As String) As String
    Dim trimmed As String
    trimmed = Trim$(textBetween)

    ' Check for semicolon
    If InStr(1, trimmed, ";") > 0 Then
        DetectSeparator = "semicolon"
        Exit Function
    End If

    ' Check for comma
    If InStr(1, trimmed, ",") > 0 Then
        DetectSeparator = "comma"
        Exit Function
    End If

    DetectSeparator = "none"
End Function

' ── Helper: check if conjunction precedes final marker ──────
Private Function DetectConjunction(ByVal textBefore As String) As String
    Dim trimmed As String
    trimmed = LCase(Trim$(textBefore))

    ' Remove trailing semicolons/commas for checking
    Do While Len(trimmed) > 0 And (Right$(trimmed, 1) = ";" Or Right$(trimmed, 1) = ",")
        trimmed = Trim$(Left$(trimmed, Len(trimmed) - 1))
    Loop

    ' Check if ends with "and" or "or"
    If Len(trimmed) >= 3 Then
        If Right$(trimmed, 4) = " and" Or trimmed = "and" Then
            DetectConjunction = "and"
            Exit Function
        End If
        If Right$(trimmed, 3) = " or" Or trimmed = "or" Then
            DetectConjunction = "or"
            Exit Function
        End If
    End If

    DetectConjunction = "none"
End Function

' ════════════════════════════════════════════════════════════
'  MAIN RULE FUNCTION
' ════════════════════════════════════════════════════════════
Public Function Check_InlineListFormat(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' Track list styles: "separator|conjunction|ending" -> count
    Dim styleCounts As New Scripting.Dictionary
    ' Store list details for flagging: Collection of Array(styleKey, paraIdx, rangeStart, rangeEnd, paraText)
    Dim listDetails As New Collection

    Dim para As Paragraph
    Dim paraIdx As Long

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        ' Page range filter
        If Not PleadingsEngine.IsInPageRange(para.Range) Then GoTo NextPara

        Dim paraText As String
        paraText = para.Range.Text

        ' Find all markers in this paragraph
        Dim markers As Collection
        Set markers = FindMarkersInPara(paraText)

        ' Need at least 2 markers to form an inline list
        If markers.Count < 2 Then GoTo NextPara

        ' Verify markers are of the same type and sequential
        Dim firstType As String
        Dim mk As Variant
        mk = markers(1)
        firstType = CStr(mk(3))

        Dim sameType As Boolean
        sameType = True
        Dim mi As Long
        For mi = 2 To markers.Count
            mk = markers(mi)
            If CStr(mk(3)) <> firstType Then
                sameType = False
                Exit For
            End If
        Next mi
        If Not sameType Then GoTo NextPara

        ' ── Analyse separator style ────────────────────────
        Dim separators As New Collection
        For mi = 2 To markers.Count
            Dim prevMk As Variant
            prevMk = markers(mi - 1)
            Dim currMk As Variant
            currMk = markers(mi)

            Dim prevEnd As Long
            prevEnd = CLng(prevMk(0)) + Len(CStr(prevMk(1)))
            Dim currStart As Long
            currStart = CLng(currMk(0))

            If currStart > prevEnd Then
                Dim between As String
                between = Mid$(paraText, prevEnd, currStart - prevEnd)
                separators.Add DetectSeparator(between)
            Else
                separators.Add "none"
            End If
        Next mi

        ' Determine dominant separator for this list
        Dim sepSemi As Long, sepComma As Long, sepNone As Long
        sepSemi = 0: sepComma = 0: sepNone = 0
        Dim s As Variant
        For Each s In separators
            Select Case CStr(s)
                Case "semicolon": sepSemi = sepSemi + 1
                Case "comma": sepComma = sepComma + 1
                Case "none": sepNone = sepNone + 1
            End Select
        Next s

        Dim listSep As String
        If sepSemi >= sepComma And sepSemi >= sepNone Then
            listSep = "semicolon"
        ElseIf sepComma >= sepSemi And sepComma >= sepNone Then
            listSep = "comma"
        Else
            listSep = "none"
        End If

        ' ── Check conjunction before final marker ──────────
        Dim lastMk As Variant
        lastMk = markers(markers.Count)
        Dim lastMkStart As Long
        lastMkStart = CLng(lastMk(0))

        Dim secondLastMk As Variant
        secondLastMk = markers(markers.Count - 1)
        Dim slEnd As Long
        slEnd = CLng(secondLastMk(0)) + Len(CStr(secondLastMk(1)))

        Dim conjText As String
        If lastMkStart > slEnd Then
            conjText = Mid$(paraText, slEnd, lastMkStart - slEnd)
        Else
            conjText = ""
        End If
        Dim conjunction As String
        conjunction = DetectConjunction(conjText)

        ' ── Check ending punctuation ───────────────────────
        Dim lastMkEnd As Long
        lastMkEnd = CLng(lastMk(0)) + Len(CStr(lastMk(1)))
        Dim afterLast As String
        If lastMkEnd <= Len(paraText) Then
            afterLast = Mid$(paraText, lastMkEnd)
        Else
            afterLast = ""
        End If
        ' Find the end of the last item (next paragraph mark or end)
        Dim ending As String
        Dim cleanAfter As String
        cleanAfter = Trim$(Replace(afterLast, vbCr, ""))
        cleanAfter = Trim$(Replace(cleanAfter, vbLf, ""))
        If Len(cleanAfter) > 0 Then
            Dim lastChar As String
            lastChar = Right$(cleanAfter, 1)
            If lastChar = "." Then
                ending = "fullstop"
            ElseIf lastChar = ";" Then
                ending = "semicolon"
            Else
                ending = "none"
            End If
        Else
            ending = "none"
        End If

        ' ── Build style key and track ──────────────────────
        Dim styleKey As String
        styleKey = listSep & "|" & conjunction & "|" & ending

        If styleCounts.Exists(styleKey) Then
            styleCounts(styleKey) = styleCounts(styleKey) + 1
        Else
            styleCounts.Add styleKey, 1
        End If

        Dim lDetail(0 To 4) As Variant
        lDetail(0) = styleKey
        lDetail(1) = paraIdx
        lDetail(2) = para.Range.Start
        lDetail(3) = para.Range.End
        lDetail(4) = Trim$(Replace(Left$(paraText, 80), vbCr, ""))
        listDetails.Add lDetail

        Set separators = Nothing
NextPara:
    Next para

    ' ── Determine dominant list style ──────────────────────
    If styleCounts.Count > 1 And listDetails.Count > 1 Then
        Dim domStyle As String
        Dim maxCnt As Long
        domStyle = ""
        maxCnt = 0
        Dim sk As Variant
        For Each sk In styleCounts.keys
            If styleCounts(sk) > maxCnt Then
                maxCnt = styleCounts(sk)
                domStyle = CStr(sk)
            End If
        Next sk

        ' ── Flag deviations ────────────────────────────────
        Dim li As Long
        For li = 1 To listDetails.Count
            Dim ld As Variant
            ld = listDetails(li)
            If CStr(ld(0)) <> domStyle Then
                Dim issue As New PleadingsIssue
                Dim rng As Range
                Set rng = doc.Range(CLng(ld(2)), CLng(ld(3)))
                Dim loc As String
                loc = PleadingsEngine.GetLocationString(rng, doc)

                ' Parse dominant style for suggestion
                Dim domParts() As String
                domParts = Split(domStyle, "|")
                Dim suggStr As String
                suggStr = "Use consistent list formatting: "
                If UBound(domParts) >= 0 Then suggStr = suggStr & domParts(0) & " separators"
                If UBound(domParts) >= 1 Then suggStr = suggStr & ", '" & domParts(1) & "' conjunction"
                If UBound(domParts) >= 2 Then suggStr = suggStr & ", " & domParts(2) & " ending"

                issue.Init RULE_NAME, loc, _
                    "Inline list format inconsistency near: '" & CStr(ld(4)) & "...'", _
                    suggStr, CLng(ld(2)), CLng(ld(3)), "possible_error"
                issues.Add issue
            End If
        Next li
    End If

    On Error GoTo 0
    Set Check_InlineListFormat = issues
End Function

' ════════════════════════════════════════════════════════════
'  STANDALONE ENTRY POINT
'  Run this macro directly from the Macros dialog (Alt+F8).
'  Checks the active document and highlights all issues found.
' ════════════════════════════════════════════════════════════
Public Sub RunInlineListFormat()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Inline List Format"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim doc As Document: Set doc = ActiveDocument
    Dim issues As Collection
    Set issues = Check_InlineListFormat(doc)

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
           vbInformation, "Inline List Format"
End Sub
