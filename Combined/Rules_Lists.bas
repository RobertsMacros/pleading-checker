Attribute VB_Name = "Rules_Lists"
' ============================================================
' Rules_Lists.bas
' Combined module for list-related proofreading rules:
'   - Rule10: Inline list format consistency (separator style,
'     conjunction usage, ending punctuation)
'   - Rule15: List punctuation consistency (ending punctuation
'     of formal list items, final-item full stop, penultimate
'     conjunction)
'
' ENGINE WIRING NOTE:
'   Both rules are dispatched under the single aggregate toggle
'   "list_rules" in PleadingsEngine.InitRuleConfig / RunAllPleadingsRules.
'   Enabling/disabling "list_rules" controls both Check_InlineListFormat
'   and Check_ListPunctuation together.
'
' Rule 10 uses LOCAL-CONTEXT grouping: only inline lists that
' are structurally close (within the same section-like region)
' are compared for consistency.  Unrelated lists in different
' sections are judged independently.
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

' -- Rule name constants ---------------------------------------
Private Const RULE_NAME_INLINE  As String = "inline_list_format"
Private Const RULE_NAME_LISTPN  As String = "list_punctuation"

' -- Marker pattern types (Rule 10) ----------------------------
Private Const MARKER_LETTER As String = "letter"   ' (a), (b), (c)
Private Const MARKER_ROMAN  As String = "roman"    ' (i), (ii), (iii)
Private Const MARKER_NUMBER As String = "number"   ' (1), (2), (3)

' Max paragraphs between inline lists to consider them related
Private Const MAX_LIST_GAP As Long = 30

' ==============================================================
'  RULE 10 - PRIVATE HELPERS
' ==============================================================

' -- Helper: check if a parenthesized marker is a clause reference --
' Returns True if the opening paren is immediately preceded by a
' digit or letter (no space), e.g. "3(4)" or "Rule 1(a)" where
' the "1(" is adjacent. These are structural references, not lists.
Private Function IsClauseRef(ByRef paraText As String, _
                              ByVal openParen As Long) As Boolean
    Dim prevCh As String
    Dim refWords As Variant
    Dim wordEnd As Long
    Dim wStart As Long
    Dim wCh As String
    Dim prevWord As String
    Dim ri As Long
    Dim conjEnd As Long
    Dim conjWord As String
    Dim cStart As Long
    Dim cc As String
    Dim scanBack As Long

    IsClauseRef = False
    If openParen <= 1 Then Exit Function

    prevCh = Mid$(paraText, openParen - 1, 1)

    ' If preceded by a digit, letter, or closing paren -- clause ref
    If (prevCh >= "0" And prevCh <= "9") Or _
       (prevCh >= "A" And prevCh <= "Z") Or _
       (prevCh >= "a" And prevCh <= "z") Or _
       prevCh = ")" Then
        IsClauseRef = True
        Exit Function
    End If

    ' If not preceded by a space, nothing more to check
    If prevCh <> " " Then Exit Function

    ' -- Check for structural reference word before the space --
    refWords = Array("paragraph", "paragraphs", "para", "paras", _
                     "section", "sections", "sect", "sects", _
                     "clause", "clauses", "cl", _
                     "article", "articles", "art", "arts", _
                     "rule", "rules", "r", _
                     "regulation", "regulations", "reg", "regs", _
                     "schedule", "schedules", "sch", _
                     "sub-paragraph", "sub-paragraphs", _
                     "sub-section", "sub-sections", _
                     "sub-clause", "sub-clauses", _
                     "part", "parts", "pt", _
                     "item", "items", "annex")

    wordEnd = openParen - 2
    If wordEnd >= 1 Then
        wStart = wordEnd
        Do While wStart >= 1
            wCh = Mid$(paraText, wStart, 1)
            If (wCh >= "A" And wCh <= "Z") Or _
               (wCh >= "a" And wCh <= "z") Or wCh = "-" Then
                wStart = wStart - 1
            Else
                Exit Do
            End If
        Loop
        wStart = wStart + 1
        If wStart <= wordEnd Then
            prevWord = LCase(Mid$(paraText, wStart, wordEnd - wStart + 1))
            For ri = LBound(refWords) To UBound(refWords)
                If prevWord = CStr(refWords(ri)) Then
                    IsClauseRef = True
                    Exit Function
                End If
            Next ri
        End If
    End If

    ' -- Check for conjunction-linked clause ref --
    ' e.g. "paragraph (1) or (2)" -- the "(2)" preceded by "or "
    conjEnd = openParen - 2
    If conjEnd >= 1 Then
        cStart = conjEnd
        Do While cStart >= 1
            cc = Mid$(paraText, cStart, 1)
            If (cc >= "A" And cc <= "Z") Or (cc >= "a" And cc <= "z") Then
                cStart = cStart - 1
            Else
                Exit Do
            End If
        Loop
        cStart = cStart + 1
        If cStart <= conjEnd Then
            conjWord = LCase(Mid$(paraText, cStart, conjEnd - cStart + 1))
            If conjWord = "and" Or conjWord = "or" Or conjWord = "to" Then
                scanBack = cStart - 1
                Do While scanBack >= 1 And Mid$(paraText, scanBack, 1) = " "
                    scanBack = scanBack - 1
                Loop
                If scanBack >= 1 And Mid$(paraText, scanBack, 1) = ")" Then
                    IsClauseRef = True
                    Exit Function
                End If
            End If
        End If
    End If
End Function

' -- Helper: detect marker type from content between parens ----
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

' -- Helper: find all inline list markers in a paragraph -------
' Returns Collection of Array(markerPos, markerText, markerContent, markerType)
Private Function FindMarkersInPara(ByVal paraText As String) As Collection
    Dim markers As New Collection
    Dim pos As Long
    Dim openParen As Long
    Dim closeParen As Long
    Dim content As String
    Dim info() As Variant
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

        If Len(mType) > 0 And Not IsClauseRef(paraText, openParen) Then
            ReDim info(0 To 3)
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

' -- Helper: detect separator before a marker ------------------
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

' -- Helper: check if conjunction precedes final marker --------
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

' -- Helper: analyse one inline list paragraph and return its style key --
'  Returns "" if the paragraph is not a valid inline list.
Private Function AnalyseInlineList(ByVal paraText As String, _
        markers As Collection) As String

    ' Need at least 2 markers to form an inline list
    If markers.Count < 2 Then
        AnalyseInlineList = ""
        Exit Function
    End If

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
    If Not sameType Then
        AnalyseInlineList = ""
        Exit Function
    End If

    ' -- Analyse separator style --------------------------
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

    ' -- Check conjunction before final marker ----------------
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

    ' -- Check ending punctuation -----------------------------
    Dim lastMkEnd As Long
    lastMkEnd = CLng(lastMk(0)) + Len(CStr(lastMk(1)))
    Dim afterLast As String
    If lastMkEnd <= Len(paraText) Then
        afterLast = Mid$(paraText, lastMkEnd)
    Else
        afterLast = ""
    End If
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

    ' -- Build style key --------------------------------------
    AnalyseInlineList = listSep & "|" & conjunction & "|" & ending
End Function

' ==============================================================
'  RULE 10 - PUBLIC FUNCTION: Check_InlineListFormat
'
'  LOCAL-CONTEXT APPROACH:
'  1. Collect all inline-list paragraphs with their style keys.
'  2. Group consecutive inline lists that are within MAX_LIST_GAP
'     paragraphs of each other into a "cluster".
'  3. Within each cluster, determine dominant style and flag
'     deviations.
' ==============================================================
Public Function Check_InlineListFormat(doc As Document) As Collection
    Dim issues As New Collection

    On Error Resume Next

    ' -- Collect all inline list paragraphs --------------------
    ' Each entry: Array(styleKey, paraIdx, rangeStart, rangeEnd, previewText)
    Dim listCap As Long
    listCap = 64
    Dim listCount As Long
    listCount = 0
    Dim lStyles() As String
    Dim lParaIdx() As Long
    Dim lStarts() As Long
    Dim lEnds() As Long
    Dim lPreviews() As String
    ReDim lStyles(0 To listCap - 1)
    ReDim lParaIdx(0 To listCap - 1)
    ReDim lStarts(0 To listCap - 1)
    ReDim lEnds(0 To listCap - 1)
    ReDim lPreviews(0 To listCap - 1)

    Dim para As Paragraph
    Dim paraIdx As Long

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        ' Page range filter
        If Not EngineIsInPageRange(para.Range) Then GoTo NextPara

        Dim paraText As String
        paraText = para.Range.Text
        If Err.Number <> 0 Then paraText = "": Err.Clear

        ' Find all markers in this paragraph
        Dim markers As Collection
        Set markers = FindMarkersInPara(paraText)

        ' Analyse as inline list
        Dim styleKey As String
        styleKey = AnalyseInlineList(paraText, markers)
        If Len(styleKey) = 0 Then GoTo NextPara

        ' Grow arrays if needed
        If listCount >= listCap Then
            listCap = listCap * 2
            ReDim Preserve lStyles(0 To listCap - 1)
            ReDim Preserve lParaIdx(0 To listCap - 1)
            ReDim Preserve lStarts(0 To listCap - 1)
            ReDim Preserve lEnds(0 To listCap - 1)
            ReDim Preserve lPreviews(0 To listCap - 1)
        End If

        lStyles(listCount) = styleKey
        lParaIdx(listCount) = paraIdx
        lStarts(listCount) = para.Range.Start
        If Err.Number <> 0 Then lStarts(listCount) = 0: Err.Clear
        lEnds(listCount) = para.Range.End
        If Err.Number <> 0 Then lEnds(listCount) = 0: Err.Clear
        lPreviews(listCount) = Trim$(Replace(Left$(paraText, 80), vbCr, ""))
        listCount = listCount + 1

NextPara:
    Next para

    If listCount < 2 Then
        On Error GoTo 0
        Set Check_InlineListFormat = issues
        Exit Function
    End If

    ' -- Group into local clusters ----------------------------
    ' A new cluster starts when the paragraph gap exceeds MAX_LIST_GAP
    Dim csCap As Long
    csCap = 16
    Dim csCount As Long
    csCount = 0
    Dim clusterStarts() As Long
    Dim clusterEnds() As Long
    ReDim clusterStarts(0 To csCap - 1)
    ReDim clusterEnds(0 To csCap - 1)

    Dim curClusterStart As Long
    curClusterStart = 0

    Dim li As Long
    For li = 1 To listCount - 1
        Dim gap As Long
        gap = lParaIdx(li) - lParaIdx(li - 1)
        If gap > MAX_LIST_GAP Then
            ' Close current cluster
            If csCount >= csCap Then
                csCap = csCap * 2
                ReDim Preserve clusterStarts(0 To csCap - 1)
                ReDim Preserve clusterEnds(0 To csCap - 1)
            End If
            clusterStarts(csCount) = curClusterStart
            clusterEnds(csCount) = li - 1
            csCount = csCount + 1
            curClusterStart = li
        End If
    Next li

    ' Close last cluster
    If csCount >= csCap Then
        csCap = csCap * 2
        ReDim Preserve clusterStarts(0 To csCap - 1)
        ReDim Preserve clusterEnds(0 To csCap - 1)
    End If
    clusterStarts(csCount) = curClusterStart
    clusterEnds(csCount) = listCount - 1
    csCount = csCount + 1

    ' -- Within each cluster, find dominant and flag -----------
    Dim ci As Long
    For ci = 0 To csCount - 1
        Dim cStart As Long
        cStart = clusterStarts(ci)
        Dim cEnd As Long
        cEnd = clusterEnds(ci)

        ' Need at least 2 lists in cluster to compare
        If cEnd - cStart < 1 Then GoTo NextCluster

        ' Count styles in this cluster
        Dim styleCounts As Object
        Set styleCounts = CreateObject("Scripting.Dictionary")
        Dim cj As Long
        For cj = cStart To cEnd
            If styleCounts.Exists(lStyles(cj)) Then
                styleCounts(lStyles(cj)) = styleCounts(lStyles(cj)) + 1
            Else
                styleCounts.Add lStyles(cj), 1
            End If
        Next cj

        ' Only flag if more than one style in this cluster
        If styleCounts.Count < 2 Then GoTo NextCluster

        ' Find dominant style
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

        ' Flag deviations
        For cj = cStart To cEnd
            If lStyles(cj) <> domStyle Then
                Dim finding As Object
                Dim rng As Range
                Set rng = doc.Range(lStarts(cj), lEnds(cj))
                If Err.Number <> 0 Then Err.Clear: GoTo NextClusterItem
                Dim loc As String
                loc = EngineGetLocationString(rng, doc)
                If Err.Number <> 0 Then loc = "unknown location": Err.Clear

                ' Parse dominant style for suggestion
                Dim domParts() As String
                domParts = Split(domStyle, "|")
                Dim suggStr As String
                suggStr = "Use consistent list formatting: "
                If UBound(domParts) >= 0 Then suggStr = suggStr & domParts(0) & " separators"
                If UBound(domParts) >= 1 Then suggStr = suggStr & ", '" & domParts(1) & "' conjunction"
                If UBound(domParts) >= 2 Then suggStr = suggStr & ", " & domParts(2) & " ending"

                Set finding = CreateIssueDict(RULE_NAME_INLINE, loc, "Inline list format inconsistency near: '" & lPreviews(cj) & "...'", suggStr, lStarts(cj), lEnds(cj), "possible_error")
                issues.Add finding
            End If
NextClusterItem:
        Next cj
NextCluster:
    Next ci

    On Error GoTo 0
    Set Check_InlineListFormat = issues
End Function

' ==============================================================
'  RULE 15 - PRIVATE HELPERS
' ==============================================================

' -- Strip trailing carriage return / line feed ----------------
Private Function StripTrailingCr(ByVal text As String) As String
    Dim result As String
    result = text

    Do While Len(result) > 0
        Dim lastCh As String
        lastCh = Right(result, 1)
        If lastCh = vbCr Or lastCh = vbLf Or lastCh = Chr(13) Or lastCh = Chr(10) Then
            result = Left(result, Len(result) - 1)
        Else
            Exit Do
        End If
    Loop

    StripTrailingCr = result
End Function

' -- Get last N characters of a string -------------------------
Private Function GetLastNChars(ByVal text As String, ByVal n As Long) As String
    If Len(text) <= n Then
        GetLastNChars = text
    Else
        GetLastNChars = Right(text, n)
    End If
End Function

' -- Classify the ending punctuation of a list item ------------
Private Function ClassifyEnding(ByVal text As String) As String
    Dim trimmed As String
    Dim endChar As String

    trimmed = StripTrailingCr(text)
    trimmed = Trim(trimmed)

    If Len(trimmed) = 0 Then
        ClassifyEnding = "none"
        Exit Function
    End If

    endChar = Right(trimmed, 1)

    Select Case endChar
        Case ";"
            ClassifyEnding = "semicolon"
        Case "."
            ClassifyEnding = "full_stop"
        Case ","
            ClassifyEnding = "comma"
        Case ":"
            ClassifyEnding = "colon"
        Case Else
            ClassifyEnding = "none"
    End Select
End Function

' -- Process a single list group for punctuation issues --------
Private Sub ProcessListGroup(doc As Document, _
                              ByRef issues As Collection, _
                              ByRef paraStarts() As Long, _
                              ByRef paraEnds() As Long, _
                              ByRef paraTexts() As String, _
                              ByVal groupStart As Long, _
                              ByVal groupEnd As Long)
    Dim itemCount As Long
    Dim i As Long
    Dim endings() As String
    Dim endingCounts As Object ' Dictionary
    Dim dominantEnding As String
    Dim maxCount As Long

    itemCount = groupEnd - groupStart + 1
    If itemCount < 2 Then Exit Sub ' Single-item list, nothing to check

    ' -- Classify the ending of each list item ------------------
    ReDim endings(groupStart To groupEnd)

    For i = groupStart To groupEnd
        endings(i) = ClassifyEnding(paraTexts(i))
    Next i

    ' -- Count endings to find dominant -------------------------
    Set endingCounts = CreateObject("Scripting.Dictionary")

    For i = groupStart To groupEnd
        If endingCounts.Exists(endings(i)) Then
            endingCounts(endings(i)) = endingCounts(endings(i)) + 1
        Else
            endingCounts.Add endings(i), 1
        End If
    Next i

    dominantEnding = ""
    maxCount = 0
    Dim key As Variant
    For Each key In endingCounts.keys
        If endingCounts(key) > maxCount Then
            maxCount = endingCounts(key)
            dominantEnding = CStr(key)
        End If
    Next key

    ' -- Flag items that deviate from dominant ending ------------
    For i = groupStart To groupEnd
        If endings(i) <> dominantEnding Then
            ' Skip the last item if dominant is semicolon (special rule below)
            If dominantEnding = "semicolon" And i = groupEnd Then
                GoTo ContinueItem
            End If

            Dim rng As Range
            Dim locStr As String
            Dim finding As Object

            On Error Resume Next
            Set rng = doc.Range(paraStarts(i), paraEnds(i))
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
                GoTo ContinueItem
            End If

            If Not EngineIsInPageRange(rng) Then
                On Error GoTo 0
                GoTo ContinueItem
            End If

            locStr = EngineGetLocationString(rng, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If
            On Error GoTo 0

            Set finding = CreateIssueDict(RULE_NAME_LISTPN, locStr, "List item ending '" & endings(i) & "' differs from " & "dominant ending '" & dominantEnding & "'", "Change ending punctuation to match list style (" & dominantEnding & ")", paraStarts(i), paraEnds(i), "possible_error")
            issues.Add finding
        End If

ContinueItem:
    Next i

    ' -- Special: if dominant is semicolon, last item should end with full stop -
    If dominantEnding = "semicolon" Then
        If endings(groupEnd) <> "full_stop" Then
            On Error Resume Next
            Set rng = doc.Range(paraStarts(groupEnd), paraEnds(groupEnd))
            If Err.Number = 0 Then
                If EngineIsInPageRange(rng) Then
                    locStr = EngineGetLocationString(rng, doc)
                    If Err.Number <> 0 Then
                        locStr = "unknown location"
                        Err.Clear
                    End If

                    Set finding = CreateIssueDict(RULE_NAME_LISTPN, locStr, "Last list item should end with a full stop, not '" & endings(groupEnd) & "'", "End the final list item with a full stop", paraStarts(groupEnd), paraEnds(groupEnd), "possible_error")
                    issues.Add finding
                End If
            End If
            On Error GoTo 0
        End If

        ' -- Check penultimate item for "and" or "or" -----------
        If itemCount >= 2 Then
            Dim penIdx As Long
            penIdx = groupEnd - 1
            Dim penText As String
            penText = LCase(Trim(StripTrailingCr(paraTexts(penIdx))))

            Dim hasConjunction As Boolean
            hasConjunction = False

            ' Check if text ends with "and;" or "or;" or similar
            If Right(penText, 4) = "and;" Or Right(penText, 3) = "or;" Or _
               Right(penText, 4) = "and," Or Right(penText, 3) = "or," Or _
               Right(penText, 3) = "and" Or Right(penText, 2) = "or" Then
                hasConjunction = True
            End If

            ' Also check for "and" / "or" as last word before punctuation
            Dim lastWords As String
            lastWords = GetLastNChars(penText, 10)
            If InStr(1, lastWords, " and") > 0 Or InStr(1, lastWords, " or") > 0 Then
                hasConjunction = True
            End If

            If Not hasConjunction Then
                On Error Resume Next
                Set rng = doc.Range(paraStarts(penIdx), paraEnds(penIdx))
                If Err.Number = 0 Then
                    If EngineIsInPageRange(rng) Then
                        locStr = EngineGetLocationString(rng, doc)
                        If Err.Number <> 0 Then
                            locStr = "unknown location"
                            Err.Clear
                        End If

                        Set finding = CreateIssueDict(RULE_NAME_LISTPN, locStr, "Penultimate list item should include 'and' or 'or' " & "before terminal punctuation", "Add 'and' or 'or' before the semicolon", paraStarts(penIdx), paraEnds(penIdx), "possible_error")
                        issues.Add finding
                    End If
                End If
                On Error GoTo 0
            End If
        End If
    End If
End Sub

' ==============================================================
'  RULE 15 - PUBLIC FUNCTION: Check_ListPunctuation
' ==============================================================
Public Function Check_ListPunctuation(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long
    Dim totalParas As Long

    ' -- Collect all paragraphs into arrays for easier processing -
    totalParas = doc.Paragraphs.Count
    If totalParas = 0 Then
        Set Check_ListPunctuation = issues
        Exit Function
    End If

    ' Arrays to hold paragraph info
    Dim paraStarts() As Long
    Dim paraEnds() As Long
    Dim paraTexts() As String
    Dim paraIsList() As Boolean
    Dim paraListID() As Long

    ReDim paraStarts(1 To totalParas)
    ReDim paraEnds(1 To totalParas)
    ReDim paraTexts(1 To totalParas)
    ReDim paraIsList(1 To totalParas)
    ReDim paraListID(1 To totalParas)

    paraIdx = 0
    On Error Resume Next
    For Each para In doc.Paragraphs
        Err.Clear
        paraIdx = paraIdx + 1

        Dim paraRange As Range
        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            paraStarts(paraIdx) = 0
            paraEnds(paraIdx) = 0
            paraTexts(paraIdx) = ""
            paraIsList(paraIdx) = False
            paraListID(paraIdx) = 0
            GoTo NextParaCollect
        End If

        paraStarts(paraIdx) = paraRange.Start
        paraEnds(paraIdx) = paraRange.End
        paraTexts(paraIdx) = paraRange.Text

        ' Check if paragraph is a list item
        Dim listType As Long
        listType = 0
        listType = paraRange.ListFormat.ListType
        If Err.Number <> 0 Then
            Err.Clear
            listType = 0
        End If

        paraIsList(paraIdx) = (listType <> 0) ' 0 = wdListNoNumbering

        ' Get a list identifier for grouping
        Dim listID As Long
        listID = 0
        If paraIsList(paraIdx) Then
            listID = paraRange.ListFormat.List.ListParagraphs.Count
            If Err.Number <> 0 Then
                Err.Clear
                ' Fallback: use list level + approximate position
                listID = paraRange.ListFormat.ListLevelNumber + 1
                If Err.Number <> 0 Then
                    Err.Clear
                    listID = 1
                End If
            End If
        End If
        paraListID(paraIdx) = listID

NextParaCollect:
    Next para
    On Error GoTo 0

    ' -- Group consecutive list paragraphs into lists -----------
    Dim groupStart As Long
    Dim groupEnd As Long
    Dim inGroup As Boolean

    inGroup = False
    Dim p As Long

    For p = 1 To totalParas
        If paraIsList(p) Then
            If Not inGroup Then
                groupStart = p
                inGroup = True
            End If
            groupEnd = p
        Else
            If inGroup Then
                ' Process the list group
                ProcessListGroup doc, issues, paraStarts, paraEnds, paraTexts, _
                                 groupStart, groupEnd
                inGroup = False
            End If
        End If
    Next p

    ' Process final group if document ends with a list
    If inGroup Then
        ProcessListGroup doc, issues, paraStarts, paraEnds, paraTexts, _
                         groupStart, groupEnd
    End If

    Set Check_ListPunctuation = issues
End Function


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
'  Late-bound wrapper: PleadingsEngine.IsInPageRange
' ----------------------------------------------------------------
Private Function EngineIsInPageRange(rng As Object) As Boolean
    On Error Resume Next
    EngineIsInPageRange = Application.Run("PleadingsEngine.IsInPageRange", rng)
    If Err.Number <> 0 Then
        Debug.Print "EngineIsInPageRange: fallback (Err " & Err.Number & ": " & Err.Description & ")"
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
        Debug.Print "EngineGetLocationString: fallback (Err " & Err.Number & ": " & Err.Description & ")"
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function
