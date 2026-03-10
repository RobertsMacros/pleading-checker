Attribute VB_Name = "Rules_Numbering"
' ============================================================
' Rules_Numbering.bas
' Combined proofreading rules for numbering:
'   - Rule03: Sequential numbering (Check_SequentialNumbering)
'   - Rule08: Clause number format (Check_ClauseNumberFormat)
'
' Dependencies:
'   - PleadingsEngine.bas (IsInPageRange, GetLocationString)
' ============================================================
Option Explicit

Private Const RULE_NAME_SEQ As String = "sequential_numbering"
Private Const RULE_NAME_FMT As String = "clause_number_format"

' ============================================================
'  RULE 03 -- MAIN ENTRY POINT
' ============================================================
Public Function Check_SequentialNumbering(doc As Document) As Collection
    Dim issues As New Collection

    ' -- Check Word-native numbered lists ------------------
    CheckNativeListNumbering doc, issues

    ' -- Check manually typed numbering --------------------
    CheckManualNumbering doc, issues

    Set Check_SequentialNumbering = issues
End Function

' ============================================================
'  PRIVATE: Check Word-native list numbering
'  Uses a Scripting.Dictionary keyed by list identifier to
'  track expected next values per list and level.
'
'  Each top-level key maps to a Dictionary of levels, where
'  each level stores the expected next value.
' ============================================================
Private Sub CheckNativeListNumbering(doc As Document, _
                                      ByRef issues As Collection)
    Dim listContexts As Object  ' listKey -> Dictionary(level -> expectedNext)
    Set listContexts = CreateObject("Scripting.Dictionary")
    Dim para As Paragraph
    Dim paraRange As Range
    Dim listType As Long
    Dim listKey As String
    Dim listLevel As Long
    Dim listValue As Long
    Dim expectedNext As Long
    Dim levelDict As Object
    Dim finding As Object
    Dim locStr As String
    Dim issueText As String
    Dim suggestion As String
    Dim prevLevel As Long

    ' Track the previous level per list to detect level changes
    Dim prevLevelDict As Object  ' listKey -> prevLevel
    Set prevLevelDict = CreateObject("Scripting.Dictionary")

    On Error Resume Next

    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextNativePara
        End If

        ' -- Skip non-list paragraphs ---------------------
        listType = paraRange.ListFormat.listType
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextNativePara
        End If

        ' wdListNoNumbering = 0; skip these
        If listType = 0 Then GoTo NextNativePara

        ' Only check numbered lists (wdListSimpleNumbering=1,
        ' wdListOutlineNumbering=4, wdListMixedNumbering=5)
        ' Skip bullet lists (wdListBullet=2, wdListPictureBullet=6)
        If listType = 2 Or listType = 6 Then GoTo NextNativePara

        ' -- Skip if outside configured page range --------
        If Not EngineIsInPageRange(paraRange) Then
            GoTo NextNativePara
        End If

        ' -- Determine list key (unique identifier) -------
        ' Try to use the List object's ListID first; fall back
        ' to a synthetic key built from type + position.
        listKey = ""
        Err.Clear
        Dim lstObj As Object
        Set lstObj = paraRange.ListFormat.List
        If Err.Number = 0 And Not lstObj Is Nothing Then
            listKey = "List_" & CStr(ObjPtr(lstObj))
        End If
        If Err.Number <> 0 Or Len(listKey) = 0 Then
            Err.Clear
            ' Synthetic key: use list type and an approximation
            listKey = "Synth_" & CStr(listType) & "_" & CStr(paraRange.ListFormat.ListLevelNumber)
        End If
        Err.Clear

        ' -- Get current list value and level -------------
        listValue = paraRange.ListFormat.listValue
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextNativePara
        End If

        listLevel = paraRange.ListFormat.ListLevelNumber
        If Err.Number <> 0 Then
            Err.Clear
            listLevel = 1
        End If

        ' -- Initialise tracking for this list if new -----
        If Not listContexts.Exists(listKey) Then
            Dim newLevelDict As Object
            Set newLevelDict = CreateObject("Scripting.Dictionary")
            listContexts.Add listKey, newLevelDict
            prevLevelDict.Add listKey, 0
        End If

        Set levelDict = listContexts(listKey)
        prevLevel = prevLevelDict(listKey)

        ' -- Handle level changes -------------------------
        ' When we go to a deeper level, that level starts fresh.
        ' When we return to a shallower level, reset all deeper levels.
        If listLevel <> prevLevel And prevLevel > 0 Then
            If listLevel < prevLevel Then
                ' Returning to shallower level: reset deeper levels
                Dim resetLevel As Variant
                Dim keysToRemove As New Collection
                For Each resetLevel In levelDict.keys
                    If CLng(resetLevel) > listLevel Then
                        keysToRemove.Add resetLevel
                    End If
                Next resetLevel
                Dim removeIdx As Long
                For removeIdx = 1 To keysToRemove.Count
                    levelDict.Remove keysToRemove(removeIdx)
                Next removeIdx
                Set keysToRemove = Nothing
            End If
        End If

        ' -- Check expected value at this level -----------
        If Not levelDict.Exists(listLevel) Then
            ' First item at this level in this list; record starting value
            levelDict.Add listLevel, listValue + 1
        Else
            expectedNext = levelDict(listLevel)

            If listValue = expectedNext Then
                ' Correct sequence; update expected next
                levelDict(listLevel) = listValue + 1

            ElseIf listValue = expectedNext - 1 Then
                ' Duplicate number
                Err.Clear
                locStr = EngineGetLocationString(paraRange, doc)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                End If

                issueText = "Duplicate number " & listValue & " at level " & listLevel
                suggestion = "Expected " & expectedNext & "; remove or renumber the duplicate"

                Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
                issues.Add finding
                ' Do not advance expectedNext for duplicates

            ElseIf listValue > expectedNext Then
                ' Skipped item(s)
                Err.Clear
                locStr = EngineGetLocationString(paraRange, doc)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                End If

                issueText = "Expected " & expectedNext & " but found " & listValue & _
                            " -- possible skipped item(s)"
                suggestion = "Check whether items " & expectedNext & " through " & _
                             (listValue - 1) & " are missing"

                Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
                issues.Add finding

                ' Update expected to continue from current
                levelDict(listLevel) = listValue + 1

            ElseIf listValue < expectedNext - 1 Then
                ' Numbering went backwards
                Err.Clear
                locStr = EngineGetLocationString(paraRange, doc)
                If Err.Number <> 0 Then
                    locStr = "unknown location"
                    Err.Clear
                End If

                issueText = "Expected " & expectedNext & " but found " & listValue & _
                            " -- numbering went backwards"
                suggestion = "Renumber this item to " & expectedNext & " or check list continuity"

                Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
                issues.Add finding

                ' Update expected to continue from current
                levelDict(listLevel) = listValue + 1
            Else
                ' Normal sequence
                levelDict(listLevel) = listValue + 1
            End If
        End If

        ' Record previous level for this list
        prevLevelDict(listKey) = listLevel

NextNativePara:
    Next para
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Check manually typed numbering
'  Detects paragraphs that start with a number pattern
'  (e.g. "1.", "2.", "12.3") but have no Word list formatting.
'  Tracks these separately and checks for sequence breaks.
' ============================================================
Private Sub CheckManualNumbering(doc As Document, _
                                  ByRef issues As Collection)
    Dim para As Paragraph
    Dim paraRange As Range
    Dim paraText As String
    Dim listType As Long
    Dim manualNum As Long
    Dim expectedNext As Long
    Dim tracking As Boolean
    Dim finding As Object
    Dim locStr As String
    Dim issueText As String
    Dim suggestion As String

    expectedNext = 0
    tracking = False

    On Error Resume Next

    For Each para In doc.Paragraphs
        Err.Clear

        Set paraRange = para.Range
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextManualPara
        End If

        paraText = Trim(paraRange.Text)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextManualPara
        End If

        ' -- Only process non-list paragraphs -------------
        listType = paraRange.ListFormat.listType
        If Err.Number <> 0 Then
            Err.Clear
            listType = 0
        End If

        ' If this paragraph has Word list formatting, skip it
        ' and break any manual tracking chain
        If listType <> 0 Then
            tracking = False
            expectedNext = 0
            GoTo NextManualPara
        End If

        ' -- Check if paragraph starts with a number pattern -
        ' Patterns: "N." or "N)" where N is one or more digits
        manualNum = ExtractLeadingNumber(paraText)

        If manualNum < 0 Then
            ' No number pattern found; break tracking chain
            ' but only if the paragraph has substantial text
            ' (skip blank lines to allow gaps between items)
            If Len(paraText) > 1 Then
                tracking = False
                expectedNext = 0
            End If
            GoTo NextManualPara
        End If

        ' -- Skip if outside configured page range --------
        If Not EngineIsInPageRange(paraRange) Then
            GoTo NextManualPara
        End If

        ' -- Start or continue tracking -------------------
        If Not tracking Then
            ' First manually numbered paragraph in a sequence
            tracking = True
            expectedNext = manualNum + 1
            GoTo NextManualPara
        End If

        ' -- Check sequence -------------------------------
        If manualNum = expectedNext Then
            ' Correct sequence
            expectedNext = manualNum + 1

        ElseIf manualNum > expectedNext Then
            ' Skipped item(s)
            Err.Clear
            locStr = EngineGetLocationString(paraRange, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If

            issueText = "Manual numbering: expected " & expectedNext & _
                        " but found " & manualNum & " -- possible skipped item(s)"
            suggestion = "Check whether items " & expectedNext & " through " & _
                         (manualNum - 1) & " are missing"

            Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
            issues.Add finding

            expectedNext = manualNum + 1

        ElseIf manualNum < expectedNext And manualNum = expectedNext - 1 Then
            ' Duplicate
            Err.Clear
            locStr = EngineGetLocationString(paraRange, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If

            issueText = "Manual numbering: duplicate number " & manualNum
            suggestion = "Remove or renumber the duplicate item"

            Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
            issues.Add finding

        ElseIf manualNum < expectedNext - 1 Then
            ' Backwards
            Err.Clear
            locStr = EngineGetLocationString(paraRange, doc)
            If Err.Number <> 0 Then
                locStr = "unknown location"
                Err.Clear
            End If

            issueText = "Manual numbering: expected " & expectedNext & _
                        " but found " & manualNum & " -- numbering went backwards"
            suggestion = "Renumber this item to " & expectedNext & " or check sequence"

            Set finding = CreateIssueDict(RULE_NAME_SEQ, locStr, issueText, suggestion, paraRange.Start, paraRange.End, "error")
            issues.Add finding

            expectedNext = manualNum + 1
        Else
            ' Normal (covers any other case)
            expectedNext = manualNum + 1
        End If

NextManualPara:
    Next para
    On Error GoTo 0
End Sub

' ============================================================
'  PRIVATE: Extract leading number from paragraph text
'  Returns the number if the text starts with a pattern like
'  "1.", "12.", "3)", "42)"; returns -1 if no match.
'  Uses the VBA Like operator for pattern matching.
' ============================================================
Private Function ExtractLeadingNumber(ByVal txt As String) As Long
    Dim trimmed As String
    Dim numStr As String
    Dim i As Long
    Dim ch As String

    trimmed = Trim(txt)
    ExtractLeadingNumber = -1

    If Len(trimmed) = 0 Then Exit Function

    ' Check first character is a digit
    If Not (trimmed Like "#*") Then Exit Function

    ' Extract consecutive digits from the start
    numStr = ""
    For i = 1 To Len(trimmed)
        ch = Mid(trimmed, i, 1)
        If ch >= "0" And ch <= "9" Then
            numStr = numStr & ch
        Else
            Exit For
        End If
    Next i

    ' Must have at least one digit
    If Len(numStr) = 0 Then Exit Function

    ' The character after the digits must be "." or ")"
    ' to qualify as a numbering pattern
    If i <= Len(trimmed) Then
        ch = Mid(trimmed, i, 1)
        If ch = "." Or ch = ")" Then
            On Error Resume Next
            ExtractLeadingNumber = CLng(numStr)
            If Err.Number <> 0 Then
                ExtractLeadingNumber = -1
                Err.Clear
            End If
            On Error GoTo 0
        End If
    End If
End Function

' ============================================================
'  PRIVATE: Extract clause number prefix from paragraph text
'  Returns the clause number prefix or empty string if none
'  found.
' ============================================================
Private Function ExtractClausePrefix(ByVal txt As String) As String
    Dim cleanText As String
    cleanText = Trim$(Replace(txt, vbCr, ""))
    cleanText = Trim$(Replace(cleanText, vbLf, ""))
    If Len(cleanText) = 0 Then
        ExtractClausePrefix = ""
        Exit Function
    End If

    ' A clause number starts at the beginning and ends before
    ' the first space or tab that is followed by non-number text
    Dim i As Long
    Dim ch As String
    Dim prefix As String
    prefix = ""

    ' Must start with a digit
    If Not (Left$(cleanText, 1) Like "[0-9]") Then
        ExtractClausePrefix = ""
        Exit Function
    End If

    ' Collect characters that form the clause number
    ' Valid clause number chars: digits, dots, parens, lowercase letters
    For i = 1 To Len(cleanText)
        ch = Mid$(cleanText, i, 1)
        If ch Like "[0-9]" Or ch = "." Or ch = "(" Or ch = ")" Or _
           (ch Like "[a-z]" And i > 1) Or (ch Like "[ivxlcdm]" And i > 1) Then
            prefix = prefix & ch
        ElseIf ch = " " Or ch = vbTab Or ch = Chr(9) Then
            ' End of clause number
            Exit For
        Else
            ' Non-clause character encountered
            Exit For
        End If
    Next i

    ' Validate: must contain at least one digit
    Dim hasDigit As Boolean
    hasDigit = False
    For i = 1 To Len(prefix)
        If Mid$(prefix, i, 1) Like "[0-9]" Then
            hasDigit = True
            Exit For
        End If
    Next i

    If Not hasDigit Then
        ExtractClausePrefix = ""
        Exit Function
    End If

    ' Remove trailing dots (e.g., "1." -> "1")
    Do While Len(prefix) > 0 And Right$(prefix, 1) = "."
        prefix = Left$(prefix, Len(prefix) - 1)
    Loop

    ExtractClausePrefix = prefix
End Function

' ============================================================
'  PRIVATE: Classify the clause number format
'  Returns a format pattern string describing the style
' ============================================================
Private Function ClassifyClauseFormat(ByVal prefix As String) As String
    If Len(prefix) = 0 Then
        ClassifyClauseFormat = ""
        Exit Function
    End If

    ' Level 1: plain number like "1" or "12"
    If prefix Like "#" Or prefix Like "##" Or prefix Like "###" Then
        ClassifyClauseFormat = "L1_plain"
        Exit Function
    End If

    ' Level 2: dotted like "1.1", "12.3", "1.12"
    If prefix Like "#.#" Or prefix Like "##.#" Or prefix Like "#.##" Or _
       prefix Like "##.##" Then
        ClassifyClauseFormat = "L2_dotted"
        Exit Function
    End If

    ' Level 3 style A: "1.1(a)" -- dotted number followed by (letter)
    If prefix Like "#.#(*)" Or prefix Like "##.#(*)" Or _
       prefix Like "#.##(*)" Or prefix Like "##.##(*)" Then
        ' Check if content in parens is a lowercase letter
        Dim parenContent As String
        Dim pOpen As Long
        pOpen = InStr(1, prefix, "(")
        If pOpen > 0 Then
            Dim pClose As Long
            pClose = InStr(pOpen, prefix, ")")
            If pClose > pOpen + 1 Then
                parenContent = Mid$(prefix, pOpen + 1, pClose - pOpen - 1)
                If Len(parenContent) = 1 And parenContent Like "[a-z]" Then
                    ClassifyClauseFormat = "L3_dotted_letter"
                    Exit Function
                End If
            End If
        End If
    End If

    ' Level 3 style B: "1.1.1" -- triple dotted
    If prefix Like "#.#.#" Or prefix Like "##.#.#" Or _
       prefix Like "#.##.#" Or prefix Like "#.#.##" Then
        ClassifyClauseFormat = "L3_dotted_sub"
        Exit Function
    End If

    ' Level 4: double parenthetical like "1.1(a)(i)"
    Dim parenCount As Long
    Dim ci As Long
    parenCount = 0
    For ci = 1 To Len(prefix)
        If Mid$(prefix, ci, 1) = "(" Then parenCount = parenCount + 1
    Next ci
    If parenCount >= 2 Then
        ClassifyClauseFormat = "L4_double_paren"
        Exit Function
    End If

    ' Single parenthetical at end: "(a)" or "(i)" style
    If Right$(prefix, 1) = ")" Then
        pOpen = InStrRev(prefix, "(")
        If pOpen > 0 Then
            parenContent = Mid$(prefix, pOpen + 1, Len(prefix) - pOpen - 1)
            If Len(parenContent) = 1 And parenContent Like "[a-z]" Then
                ClassifyClauseFormat = "L3_paren_letter"
                Exit Function
            End If
            ' Roman numeral in parens
            Dim allRoman As Boolean
            allRoman = True
            For ci = 1 To Len(parenContent)
                If Not (Mid$(parenContent, ci, 1) Like "[ivxlcdm]") Then
                    allRoman = False
                    Exit For
                End If
            Next ci
            If allRoman And Len(parenContent) > 0 Then
                ClassifyClauseFormat = "L3_paren_roman"
                Exit Function
            End If
        End If
    End If

    ' Fallback: generic numbered
    ClassifyClauseFormat = "other_" & prefix
End Function

' ============================================================
'  RULE 08 -- MAIN ENTRY POINT
' ============================================================
Public Function Check_ClauseNumberFormat(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long

    On Error Resume Next

    ' Track format patterns: formatPattern -> Collection of Array(paraIdx, prefix, rangeStart, rangeEnd)
    Dim formatCounts As Object
    Set formatCounts = CreateObject("Scripting.Dictionary")
    Dim clauseInfos As New Collection
    Dim cInfo() As Variant

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        ' Skip headings (they have their own numbering rules)
        If para.OutlineLevel >= wdOutlineLevel1 And _
           para.OutlineLevel <= wdOutlineLevel9 Then GoTo NextClausePara

        ' Page range filter
        If Not EngineIsInPageRange(para.Range) Then GoTo NextClausePara

        ' Extract clause number prefix
        Dim prefix As String
        prefix = ExtractClausePrefix(para.Range.Text)
        If Len(prefix) = 0 Then GoTo NextClausePara

        ' Classify the format
        Dim fmt As String
        fmt = ClassifyClauseFormat(prefix)
        If Len(fmt) = 0 Then GoTo NextClausePara

        ' Count format occurrences
        If formatCounts.Exists(fmt) Then
            formatCounts(fmt) = formatCounts(fmt) + 1
        Else
            formatCounts.Add fmt, 1
        End If

        ' Store clause info
        ReDim cInfo(0 To 3)
        cInfo(0) = paraIdx
        cInfo(1) = prefix
        cInfo(2) = para.Range.Start
        cInfo(3) = para.Range.End
        clauseInfos.Add Array(fmt, cInfo)
NextClausePara:
    Next para

    ' -- Group by level category and detect mixed formats ----
    ' Level categories: L1, L2, L3, L4
    Dim levelGroups As Object  ' "L1" -> Dictionary(format -> count)
    Set levelGroups = CreateObject("Scripting.Dictionary")
    Dim fk As Variant
    For Each fk In formatCounts.keys
        Dim levelCat As String
        If Left$(CStr(fk), 2) = "L1" Then
            levelCat = "L1"
        ElseIf Left$(CStr(fk), 2) = "L2" Then
            levelCat = "L2"
        ElseIf Left$(CStr(fk), 2) = "L3" Then
            levelCat = "L3"
        ElseIf Left$(CStr(fk), 2) = "L4" Then
            levelCat = "L4"
        Else
            levelCat = "other"
        End If

        If Not levelGroups.Exists(levelCat) Then
            levelGroups.Add levelCat, CreateObject("Scripting.Dictionary")
        End If
        Dim lgDict As Object
        Set lgDict = levelGroups(levelCat)
        lgDict.Add CStr(fk), formatCounts(fk)
    Next fk

    ' -- Find dominant format per level and flag deviations --
    Dim dominantFormats As Object  ' levelCat -> dominant format string
    Set dominantFormats = CreateObject("Scripting.Dictionary")
    Dim lgKey As Variant
    For Each lgKey In levelGroups.keys
        Set lgDict = levelGroups(lgKey)
        If lgDict.Count > 1 Then
            ' Mixed formats at this level -- find dominant
            Dim domFmt As String
            Dim maxCnt As Long
            domFmt = ""
            maxCnt = 0
            For Each fk In lgDict.keys
                If lgDict(fk) > maxCnt Then
                    maxCnt = lgDict(fk)
                    domFmt = CStr(fk)
                End If
            Next fk
            dominantFormats.Add CStr(lgKey), domFmt
        End If
    Next lgKey

    ' -- Flag individual clauses that deviate ----------------
    If dominantFormats.Count > 0 Then
        Dim ci As Long
        For ci = 1 To clauseInfos.Count
            Dim clauseArr As Variant
            clauseArr = clauseInfos(ci)
            Dim clauseFmt As String
            clauseFmt = CStr(clauseArr(0))
            Dim clauseData As Variant
            clauseData = clauseArr(1)

            ' Determine level category
            If Left$(clauseFmt, 2) = "L1" Then
                levelCat = "L1"
            ElseIf Left$(clauseFmt, 2) = "L2" Then
                levelCat = "L2"
            ElseIf Left$(clauseFmt, 2) = "L3" Then
                levelCat = "L3"
            ElseIf Left$(clauseFmt, 2) = "L4" Then
                levelCat = "L4"
            Else
                levelCat = "other"
            End If

            If dominantFormats.Exists(levelCat) Then
                If clauseFmt <> dominantFormats(levelCat) Then
                    Dim finding As Object
                    Dim rng As Range
                    Set rng = doc.Range(CLng(clauseData(2)), CLng(clauseData(3)))
                    Dim loc As String
                    loc = EngineGetLocationString(rng, doc)

                    Set finding = CreateIssueDict(RULE_NAME_FMT, loc, "Mixed clause number format: '" & CStr(clauseData(1)) & "' uses style " & clauseFmt & " but dominant " & levelCat & " style is " & dominantFormats(levelCat), "Reformat to match the dominant clause numbering style", CLng(clauseData(2)), CLng(clauseData(3)), "error")
                    issues.Add finding
                End If
            End If
        Next ci
    End If

    On Error GoTo 0
    Set Check_ClauseNumberFormat = issues
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
        EngineGetLocationString = "unknown location"
        Err.Clear
    End If
    On Error GoTo 0
End Function
