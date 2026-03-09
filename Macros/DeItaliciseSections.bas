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
' Optimised for large documents (100+ pages): uses string-based
' word parsing to minimise COM round-trips, disables screen
' updating during execution.
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
'  Helper: check whether a range is (at least partly) italic.
'  Checks the first character to resolve wdUndefined ranges.
' ────────────────────────────────────────────────────────────
Private Function IsRangeItalic(doc As Document, _
                                rStart As Long, _
                                rEnd As Long) As Boolean
    On Error Resume Next
    Dim rng As Range
    Set rng = doc.Range(rStart, rEnd)
    If rng.Font.Italic = True Then
        IsRangeItalic = True
    ElseIf rng.Font.Italic = wdUndefined Then
        ' Mixed — check first character only
        Set rng = doc.Range(rStart, rStart + 1)
        IsRangeItalic = (rng.Font.Italic = True)
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
'  Core: given a trigger-word match that is already confirmed
'  italic, determine the full reference span by parsing the
'  paragraph text as a VBA string (no per-character COM calls),
'  then de-italicise.  Returns True on success.
' ────────────────────────────────────────────────────────────
Private Function DeItaliciseSpan(doc As Document, _
                                  matchStart As Long, _
                                  matchEnd As Long) As Boolean
    DeItaliciseSpan = False

    ' ── Get paragraph boundaries (single COM call) ───────────
    Dim paraRng As Range
    Set paraRng = doc.Range(matchStart, matchEnd).Paragraphs(1).Range
    Dim paraStart As Long: paraStart = paraRng.Start
    Dim paraEnd   As Long: paraEnd = paraRng.End

    ' ── Read the rest of the paragraph after the trigger word
    '    as a plain VBA string — all word parsing happens here
    '    with zero COM calls ──────────────────────────────────
    If matchEnd >= paraEnd Then Exit Function

    Dim tailRng As Range
    Set tailRng = doc.Range(matchEnd, paraEnd)
    Dim tail As String
    tail = tailRng.Text

    ' Offset within `tail`: 1-based VBA string position
    Dim pos As Long: pos = 1
    Dim tailLen As Long: tailLen = Len(tail)

    ' spanEnd tracks the document position of the last included word
    Dim spanEnd As Long: spanEnd = matchEnd
    Dim ch As String

    Do While pos <= tailLen
        ' ── Skip spaces / non-breaking spaces ────────────────
        Do While pos <= tailLen
            ch = Mid$(tail, pos, 1)
            If ch <> " " And ch <> Chr$(160) Then Exit Do
            pos = pos + 1
        Loop
        If pos > tailLen Then Exit Do

        ' ── Gather a word (contiguous non-space run) ─────────
        Dim wordPos As Long: wordPos = pos
        Do While pos <= tailLen
            ch = Mid$(tail, pos, 1)
            If ch = " " Or ch = Chr$(160) Or ch = vbCr _
               Or ch = Chr$(13) Or ch = Chr$(7) Then Exit Do
            pos = pos + 1
        Loop

        Dim word As String
        word = Mid$(tail, wordPos, pos - wordPos)
        If Len(word) = 0 Then Exit Do

        ' ── Translate string offset to document position ─────
        Dim wordDocStart As Long: wordDocStart = matchEnd + wordPos - 1
        Dim wordDocEnd   As Long: wordDocEnd = matchEnd + pos - 1

        ' ── Stop if this word is no longer italic (1 COM call
        '    per word, not per character) ─────────────────────
        If Not IsRangeItalic(doc, wordDocStart, wordDocEnd) Then Exit Do

        ' ── Stop BEFORE a word with more than 3 alpha chars ──
        If AlphaCount(word) > 3 Then Exit Do

        ' Include this word
        spanEnd = wordDocEnd
    Loop

    ' ── Block-quote guard ────────────────────────────────────
    If IsInItalicBlock(doc, matchStart, spanEnd) Then Exit Function

    ' ── De-italicise (single COM call) ───────────────────────
    Dim target As Range
    Set target = doc.Range(matchStart, spanEnd)
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

    ' ── Suppress screen redraws for speed ────────────────────
    Application.ScreenUpdating = False

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

            ' Post-find italic check (1 COM call)
            If IsRangeItalic(doc, rng.Start, rng.End) Then
                If DeItaliciseSpan(doc, rng.Start, rng.End) Then
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

            If IsRangeItalic(doc, rng2.Start, rng2.End) Then
                If Not IsInItalicBlock(doc, rng2.Start, rng2.End) Then
                    Dim target As Range
                    Set target = doc.Range(rng2.Start, rng2.End)
                    target.Font.Italic = False
                    hitCount = hitCount + 1
                End If
            End If

            searchStart = rng2.End
            If searchStart >= doc.Content.End Then Exit Do
        Loop

    Next p

    ' ── Restore state ────────────────────────────────────────
    doc.TrackRevisions = origTrack
    Application.ScreenUpdating = True

    MsgBox "Done. De-italicised " & hitCount & " item(s).", _
           vbInformation, "De-italicise Section References"
End Sub
