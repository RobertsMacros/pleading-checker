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
' Optimised for large documents (100+ pages): uses native italic
' Find filter, Duplicate/Collapse range pattern, cached document
' boundaries, string-based word parsing, and batched italic
' checks to minimise COM round-trips.
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
'  Helper: block-quote guard.  Returns True when both the
'  20 chars before AND the 20 chars after the span are
'  entirely italic — indicating a larger italic block.
'  Accepts pre-cached document boundaries to avoid
'  recalculating doc.Content.Start/End on every call.
' ────────────────────────────────────────────────────────────
Private Function IsInItalicBlock(doc As Document, _
                                  spanStart As Long, _
                                  spanEnd As Long, _
                                  docStart As Long, _
                                  docEnd As Long) As Boolean
    On Error Resume Next
    Const PAD As Long = 20

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
'  italic (by Find's format filter), determine the full
'  reference span using string-based word parsing, then
'  de-italicise.  Returns True on success.
'
'  Word parsing is pure VBA string ops (zero COM calls).
'  Italic validation uses a batched check first: if the
'  whole candidate span is italic (single COM call), we skip
'  per-word checks entirely.  Only falls back to per-word
'  on mixed-format spans (wdUndefined).
' ────────────────────────────────────────────────────────────
Private Function DeItaliciseSpan(doc As Document, _
                                  matchStart As Long, _
                                  matchEnd As Long, _
                                  docStart As Long, _
                                  docEnd As Long) As Boolean
    DeItaliciseSpan = False

    ' ── Get paragraph boundary (single COM call) ─────────────
    Dim paraRng As Range
    Set paraRng = doc.Range(matchStart, matchEnd).Paragraphs(1).Range
    Dim paraEnd As Long: paraEnd = paraRng.End

    ' ── Read the rest of the paragraph as a VBA string ───────
    If matchEnd >= paraEnd Then Exit Function

    Dim tailRng As Range
    Set tailRng = doc.Range(matchEnd, paraEnd)
    Dim tail As String
    tail = tailRng.Text

    ' ── Parse words from the string (zero COM calls) ─────────
    Dim pos As Long: pos = 1
    Dim tailLen As Long: tailLen = Len(tail)
    Dim ch As String
    Dim lastGoodOffset As Long: lastGoodOffset = 0  ' offset into tail of last included word-end

    ' We collect word boundaries (as offsets into tail) first,
    ' then do a batched italic check.
    Dim wordStarts() As Long
    Dim wordEnds() As Long
    Dim wordTexts() As String
    Dim wordCount As Long: wordCount = 0
    ReDim wordStarts(0 To 31)
    ReDim wordEnds(0 To 31)
    ReDim wordTexts(0 To 31)

    Do While pos <= tailLen
        ' Skip spaces / non-breaking spaces
        Do While pos <= tailLen
            ch = Mid$(tail, pos, 1)
            If ch <> " " And ch <> Chr$(160) Then Exit Do
            pos = pos + 1
        Loop
        If pos > tailLen Then Exit Do

        ' Gather a word (contiguous non-space run)
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

        ' Stop BEFORE a word with more than 3 alpha characters
        If AlphaCount(word) > 3 Then Exit Do

        ' Store this word for italic checking
        If wordCount > UBound(wordStarts) Then
            ReDim Preserve wordStarts(0 To wordCount * 2)
            ReDim Preserve wordEnds(0 To wordCount * 2)
            ReDim Preserve wordTexts(0 To wordCount * 2)
        End If
        wordStarts(wordCount) = wordPos
        wordEnds(wordCount) = pos
        wordTexts(wordCount) = word
        wordCount = wordCount + 1
    Loop

    ' ── Determine span end via batched italic check ──────────
    Dim spanEnd As Long: spanEnd = matchEnd

    If wordCount > 0 Then
        ' Candidate span: from matchEnd to end of last parsed word
        Dim candidateEnd As Long
        candidateEnd = matchEnd + wordEnds(wordCount - 1) - 1

        Dim candidateRng As Range
        Set candidateRng = doc.Range(matchEnd, candidateEnd)
        Dim italicState As Long
        italicState = candidateRng.Font.Italic

        If italicState = True Then
            ' Entire candidate span is italic — include all words
            spanEnd = candidateEnd
        ElseIf italicState = wdUndefined Then
            ' Mixed formatting — check per-word (still few COM calls)
            Dim w As Long
            For w = 0 To wordCount - 1
                Dim wDocStart As Long: wDocStart = matchEnd + wordStarts(w) - 1
                Dim wDocEnd   As Long: wDocEnd = matchEnd + wordEnds(w) - 1
                Dim wRng As Range
                Set wRng = doc.Range(wDocStart, wDocEnd)
                If wRng.Font.Italic = True Then
                    spanEnd = wDocEnd
                ElseIf wRng.Font.Italic = wdUndefined Then
                    ' Check first char
                    Set wRng = doc.Range(wDocStart, wDocStart + 1)
                    If wRng.Font.Italic = True Then
                        spanEnd = wDocEnd
                    Else
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next w
        End If
        ' If italicState = False, spanEnd stays at matchEnd (trigger only)
    End If

    ' ── Block-quote guard ────────────────────────────────────
    If IsInItalicBlock(doc, matchStart, spanEnd, docStart, docEnd) Then Exit Function

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

    ' ── Cache document boundaries once ───────────────────────
    Dim docStart As Long: docStart = doc.Content.Start
    Dim docEnd   As Long: docEnd = doc.Content.End

    Dim hitCount As Long
    hitCount = 0

    ' ==========================================================
    '  PART 1 — Trigger-word references
    '  Uses Find with .Font.Italic = True so Word's native
    '  search engine only returns italic matches (skips 80-90%
    '  of non-italic text).  Duplicate/Collapse pattern avoids
    '  recreating Range and Find objects each iteration.
    ' ==========================================================
    Dim triggers As Variant
    triggers = Array("Section", "Regulation", "Article", "Paragraph")

    Dim t As Long
    For t = LBound(triggers) To UBound(triggers)

        Dim rng As Range
        Set rng = doc.Content.Duplicate

        With rng.Find
            .ClearFormatting
            .Font.Italic = True
            .Replacement.ClearFormatting
            .Text = CStr(triggers(t))
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
        End With

        Do While rng.Find.Execute
            If DeItaliciseSpan(doc, rng.Start, rng.End, docStart, docEnd) Then
                hitCount = hitCount + 1
            End If
            rng.Collapse wdCollapseEnd
        Loop

    Next t

    ' ==========================================================
    '  PART 2 — Specific statutory phrases (longest first)
    '  Same optimised Find pattern with italic filter.
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

        Dim rng2 As Range
        Set rng2 = doc.Content.Duplicate

        With rng2.Find
            .ClearFormatting
            .Font.Italic = True
            .Replacement.ClearFormatting
            .Text = CStr(phrases(p))
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
        End With

        Do While rng2.Find.Execute
            If Not IsInItalicBlock(doc, rng2.Start, rng2.End, docStart, docEnd) Then
                Dim target As Range
                Set target = doc.Range(rng2.Start, rng2.End)
                target.Font.Italic = False
                hitCount = hitCount + 1
            End If
            rng2.Collapse wdCollapseEnd
        Loop

    Next p

    ' ── Restore state ────────────────────────────────────────
    doc.TrackRevisions = origTrack
    Application.ScreenUpdating = True

    MsgBox "Done. De-italicised " & hitCount & " item(s).", _
           vbInformation, "De-italicise Section References"
End Sub
