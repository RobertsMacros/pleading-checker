Attribute VB_Name = "Rule08_ClauseNumberFormat"
' ============================================================
' Rule08_ClauseNumberFormat.bas
' Validates clause numbering format consistency across the
' document. Checks that clause references follow a consistent
' pattern and flags mixed formatting styles.
' ============================================================
Option Explicit

Private Const RULE_NAME As String = "clause_number_format"

' ── Helper: extract clause number prefix from paragraph text ─
' Returns the clause number prefix or empty string if none found
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

' ── Helper: classify the clause number format ───────────────
' Returns a format pattern string describing the style
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

    ' Level 3 style A: "1.1(a)" — dotted number followed by (letter)
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

    ' Level 3 style B: "1.1.1" — triple dotted
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

' ════════════════════════════════════════════════════════════
'  MAIN RULE FUNCTION
' ════════════════════════════════════════════════════════════
Public Function Check_ClauseNumberFormat(doc As Document) As Collection
    Dim issues As New Collection
    Dim para As Paragraph
    Dim paraIdx As Long

    On Error Resume Next

    ' Track format patterns: formatPattern -> Collection of Array(paraIdx, prefix, rangeStart, rangeEnd)
    Dim formatCounts As New Scripting.Dictionary
    Dim clauseInfos As New Collection

    paraIdx = 0
    For Each para In doc.Paragraphs
        paraIdx = paraIdx + 1

        ' Skip headings (they have their own numbering rules)
        If para.OutlineLevel >= wdOutlineLevel1 And _
           para.OutlineLevel <= wdOutlineLevel9 Then GoTo NextPara

        ' Page range filter
        If Not PleadingsEngine.IsInPageRange(para.Range) Then GoTo NextPara

        ' Extract clause number prefix
        Dim prefix As String
        prefix = ExtractClausePrefix(para.Range.Text)
        If Len(prefix) = 0 Then GoTo NextPara

        ' Classify the format
        Dim fmt As String
        fmt = ClassifyClauseFormat(prefix)
        If Len(fmt) = 0 Then GoTo NextPara

        ' Count format occurrences
        If formatCounts.Exists(fmt) Then
            formatCounts(fmt) = formatCounts(fmt) + 1
        Else
            formatCounts.Add fmt, 1
        End If

        ' Store clause info
        Dim cInfo(0 To 3) As Variant
        cInfo(0) = paraIdx
        cInfo(1) = prefix
        cInfo(2) = para.Range.Start
        cInfo(3) = para.Range.End
        clauseInfos.Add Array(fmt, cInfo)
NextPara:
    Next para

    ' ── Group by level category and detect mixed formats ────
    ' Level categories: L1, L2, L3, L4
    Dim levelGroups As New Scripting.Dictionary  ' "L1" -> Dictionary(format -> count)
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
            levelGroups.Add levelCat, New Scripting.Dictionary
        End If
        Dim lgDict As Scripting.Dictionary
        Set lgDict = levelGroups(levelCat)
        lgDict.Add CStr(fk), formatCounts(fk)
    Next fk

    ' ── Find dominant format per level and flag deviations ──
    Dim dominantFormats As New Scripting.Dictionary  ' levelCat -> dominant format string
    Dim lgKey As Variant
    For Each lgKey In levelGroups.keys
        Set lgDict = levelGroups(lgKey)
        If lgDict.Count > 1 Then
            ' Mixed formats at this level — find dominant
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

    ' ── Flag individual clauses that deviate ────────────────
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
                    Dim issue As New PleadingsIssue
                    Dim rng As Range
                    Set rng = doc.Range(CLng(clauseData(2)), CLng(clauseData(3)))
                    Dim loc As String
                    loc = PleadingsEngine.GetLocationString(rng, doc)

                    issue.Init RULE_NAME, loc, _
                        "Mixed clause number format: '" & CStr(clauseData(1)) & _
                        "' uses style " & clauseFmt & " but dominant " & _
                        levelCat & " style is " & dominantFormats(levelCat), _
                        "Reformat to match the dominant clause numbering style", _
                        CLng(clauseData(2)), CLng(clauseData(3)), "error"
                    issues.Add issue
                End If
            End If
        Next ci
    End If

    On Error GoTo 0
    Set Check_ClauseNumberFormat = issues
End Function
