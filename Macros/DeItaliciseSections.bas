Attribute VB_Name = "DeItaliciseSections"
' ============================================================
' DeItaliciseSections.bas
' Standalone macro: removes italic formatting from legislative
' references and specific statutory instrument names, using
' tracked changes (enabled automatically even if off).
'
' Trigger words (case-insensitive):
'   Section, Regulation, Article, Paragraph
'
' The span runs from the trigger word forward through all words
' with three or fewer alphabetic characters (numbers,
' parentheticals, "of", "the", "Act" etc.), stopping before the
' first word with more than three alphabetic characters.
'
' Block-quote guard: if the 20 characters either side of the
' match are all italic the match is skipped (likely a quotation).
'
' Also de-italicises these specific phrases wherever they appear
' in italics (not in block quotes):
'   Capital Adequacy Regulations
'   FI (Amendment) Act
'   Liquidity Regulations
'   Bank of Uganda Act
'   Bank of Uganda Act, 1966
'
' Usage: run DeItaliciseSectionReferences from the Macros dialog
'        or assign to a keyboard shortcut / QAT button.
'
' No dependencies — fully standalone.
' ============================================================
Option Explicit

' ────────────────────────────────────────────────────────────
'  Helper: count alphabetic characters in a string
' ────────────────────────────────────────────────────────────
Private Function AlphaCount(ByVal s As String) As Long
    Dim i As Long
    Dim n As Long
    Dim c As Long
    n = 0
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Then
            n = n + 1
        End If
    Next i
    AlphaCount = n
End Function

' ────────────────────────────────────────────────────────────
'  Helper: check whether a single-character range is italic
' ────────────────────────────────────────────────────────────
Private Function IsCharItalic(doc As Document, pos As Long) As Boolean
    On Error Resume Next
    Dim cr As Range
    Set cr = doc.Range(pos, pos + 1)
    IsCharItalic = (cr.Font.Italic = True)
    On Error GoTo 0
End Function

' ────────────────────────────────────────────────────────────
'  Helper: check whether a range is (at least partly) italic.
'  Checks the first character to resolve wdUndefined ranges.
' ────────────────────────────────────────────────────────────
Private Function IsRangeItalic(rng As Range) As Boolean
    On Error Resume Next
    If rng.Font.Italic = True Then
        IsRangeItalic = True
    ElseIf rng.Font.Italic = wdUndefined Then
        IsRangeItalic = IsCharItalic(rng.Document, rng.Start)
    Else
        IsRangeItalic = False
    End If
    On Error GoTo 0
End Function

' ────────────────────────────────────────────────────────────
'  Helper: block-quote guard.  Returns True when both the
'  20 chars before AND the 20 chars after the span are
'  entirely italic — indicating a larger italic block.
' ────────────────────────────────────────────────────────────
Private Function IsInItalicBlock(doc As Document, _
                                  spanStart As Long, _
                                  spanEnd As Long) As Boolean
    On Error Resume Next
    Const PAD As Long = 20

    Dim docStart As Long: docStart = doc.Content.Start
    Dim docEnd  As Long:  docEnd = doc.Content.End

    ' ── Check text BEFORE the span ───────────────────────────
    Dim befStart As Long
    befStart = spanStart - PAD
    If befStart < docStart Then befStart = docStart

    If spanStart - befStart < 2 Then
        IsInItalicBlock = False: Exit Function
    End If

    Dim befRng As Range
    Set befRng = doc.Range(befStart, spanStart)
    If Not (befRng.Font.Italic = True) Then
        IsInItalicBlock = False: Exit Function
    End If

    ' ── Check text AFTER the span ────────────────────────────
    Dim aftEnd As Long
    aftEnd = spanEnd + PAD
    If aftEnd > docEnd Then aftEnd = docEnd

    If aftEnd - spanEnd < 2 Then
        IsInItalicBlock = False: Exit Function
    End If

    Dim aftRng As Range
    Set aftRng = doc.Range(spanEnd, aftEnd)
    If Not (aftRng.Font.Italic = True) Then
        IsInItalicBlock = False: Exit Function
    End If

    IsInItalicBlock = True
    On Error GoTo 0
End Function

' ────────────────────────────────────────────────────────────
'  Core: given a range covering a trigger word that is already
'  confirmed italic, scan forward to find the full legislative
'  reference and de-italicise it.  Returns True on success.
' ────────────────────────────────────────────────────────────
Private Function DeItaliciseSpan(doc As Document, _
                                  matchStart As Long, _
                                  matchEnd As Long) As Boolean
    DeItaliciseSpan = False

    Dim spanStart As Long: spanStart = matchStart
    Dim spanEnd   As Long: spanEnd = matchEnd

    ' Paragraph boundary — do not scan beyond it
    Dim paraRng As Range
    Set paraRng = doc.Range(matchStart, matchEnd)
    Dim paraEnd As Long
    paraEnd = paraRng.Paragraphs(1).Range.End

    Dim scanPos   As Long: scanPos = matchEnd
    Dim ch        As String
    Dim wordStart As Long
    Dim wordEnd   As Long
    Dim wordRng   As Range
    Dim trimmed   As String

    Do While scanPos < paraEnd
        ' Skip spaces / non-breaking spaces
        Do While scanPos < paraEnd
            ch = doc.Range(scanPos, scanPos + 1).Text
            If ch <> " " And ch <> Chr$(160) Then Exit Do
            scanPos = scanPos + 1
        Loop
        If scanPos >= paraEnd Then Exit Do

        ' Gather a word (contiguous non-space run)
        wordStart = scanPos
        Do While scanPos < paraEnd
            ch = doc.Range(scanPos, scanPos + 1).Text
            If ch = " " Or ch = Chr$(160) Or ch = vbCr _
               Or ch = Chr$(13) Or ch = Chr$(7) Then Exit Do
            scanPos = scanPos + 1
        Loop
        wordEnd = scanPos

        Set wordRng = doc.Range(wordStart, wordEnd)
        trimmed = Trim$(wordRng.Text)
        If Len(trimmed) = 0 Then Exit Do

        ' Stop if the next word is no longer italic
        If Not IsRangeItalic(wordRng) Then Exit Do

        ' Stop before a word with more than 3 alpha characters
        If AlphaCount(trimmed) > 3 Then Exit Do

        ' Include this word
        spanEnd = wordEnd
    Loop

    ' ── Block-quote guard ────────────────────────────────────
    If IsInItalicBlock(doc, spanStart, spanEnd) Then Exit Function

    ' ── De-italicise ─────────────────────────────────────────
    Dim target As Range
    Set target = doc.Range(spanStart, spanEnd)
    If target.Font.Italic = True Or target.Font.Italic = wdUndefined Then
        target.Font.Italic = False
        DeItaliciseSpan = True
    End If
End Function

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Sub DeItaliciseSectionReferences()
    Dim doc As Document
    Set doc = ActiveDocument

    ' ── Enable tracked changes (restore original state later) ─
    Dim origTrack As Boolean
    origTrack = doc.TrackRevisions
    doc.TrackRevisions = True

    Dim hitCount As Long
    hitCount = 0

    ' ==========================================================
    '  PART 1 — Trigger-word references
    ' ==========================================================
    Dim triggers As Variant
    triggers = Array("Section", "Regulation", "Article", "Paragraph")

    Dim t As Long
    For t = LBound(triggers) To UBound(triggers)

        Dim searchStart As Long
        searchStart = doc.Content.Start

        Do
            ' ── Plain-text Find (no formatting filter) ───────
            Dim rng As Range
            Set rng = doc.Range(searchStart, doc.Content.End)

            With rng.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = CStr(triggers(t))
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
            End With

            If Not rng.Find.Execute Then Exit Do

            ' ── Post-find italic check ───────────────────────
            If IsRangeItalic(rng) Then
                Dim mStart As Long: mStart = rng.Start
                Dim mEnd   As Long: mEnd = rng.End

                If DeItaliciseSpan(doc, mStart, mEnd) Then
                    hitCount = hitCount + 1
                End If
            End If

            ' Advance past this match
            searchStart = rng.End
            If searchStart >= doc.Content.End Then Exit Do
        Loop

    Next t

    ' ==========================================================
    '  PART 2 — Specific statutory phrases (longest first)
    ' ==========================================================
    Dim phrases As Variant
    phrases = Array( _
        "Bank of Uganda Act, 1966", _
        "Capital Adequacy Regulations", _
        "FI (Amendment) Act", _
        "Liquidity Regulations", _
        "Bank of Uganda Act")

    Dim p As Long
    For p = LBound(phrases) To UBound(phrases)

        searchStart = doc.Content.Start

        Do
            Dim rng2 As Range
            Set rng2 = doc.Range(searchStart, doc.Content.End)

            With rng2.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = CStr(phrases(p))
                .Forward = True
                .Wrap = wdFindStop
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
            End With

            If Not rng2.Find.Execute Then Exit Do

            ' Only act if italic and not in a block quote
            If IsRangeItalic(rng2) Then
                If Not IsInItalicBlock(doc, rng2.Start, rng2.End) Then
                    rng2.Font.Italic = False
                    hitCount = hitCount + 1
                End If
            End If

            searchStart = rng2.End
            If searchStart >= doc.Content.End Then Exit Do
        Loop

    Next p

    ' ── Restore original Track Changes setting ───────────────
    doc.TrackRevisions = origTrack

    MsgBox "Done. De-italicised " & hitCount & " item(s).", _
           vbInformation, "De-italicise Section References"
End Sub
