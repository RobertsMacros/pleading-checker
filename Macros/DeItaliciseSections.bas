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
'  Helper: check whether a range is (at least partly) italic
' ────────────────────────────────────────────────────────────
Private Function IsItalic(rng As Range) As Boolean
    On Error Resume Next
    If rng.Font.Italic = True Then
        IsItalic = True
    ElseIf rng.Font.Italic = wdUndefined Then
        Dim cr As Range
        Set cr = rng.Document.Range(rng.Start, rng.Start + 1)
        IsItalic = (cr.Font.Italic = True)
    Else
        IsItalic = False
    End If
    On Error GoTo 0
End Function

' ────────────────────────────────────────────────────────────
'  Helper: check whether surrounding context is all italic
'  (block-quote guard). Returns True if both the 20 chars
'  before AND the 20 chars after the span are entirely italic,
'  meaning the match sits inside a larger italic block.
' ────────────────────────────────────────────────────────────
Private Function IsInItalicBlock(doc As Document, _
                                  spanStart As Long, _
                                  spanEnd As Long) As Boolean
    On Error Resume Next
    Const PAD As Long = 20

    Dim docStart As Long
    docStart = doc.Content.Start
    Dim docEnd As Long
    docEnd = doc.Content.End

    ' ── Check text BEFORE the span ───────────────────────────
    Dim befStart As Long
    befStart = spanStart - PAD
    If befStart < docStart Then befStart = docStart

    ' If there are fewer than 2 characters before, we cannot
    ' confidently call it a block quote — treat as non-block
    If spanStart - befStart < 2 Then
        IsInItalicBlock = False
        Exit Function
    End If

    Dim befRng As Range
    Set befRng = doc.Range(befStart, spanStart)
    If befRng.Font.Italic <> True Then
        ' Not all italic before → not a block quote
        IsInItalicBlock = False
        Exit Function
    End If

    ' ── Check text AFTER the span ────────────────────────────
    Dim aftEnd As Long
    aftEnd = spanEnd + PAD
    If aftEnd > docEnd Then aftEnd = docEnd

    If aftEnd - spanEnd < 2 Then
        IsInItalicBlock = False
        Exit Function
    End If

    Dim aftRng As Range
    Set aftRng = doc.Range(spanEnd, aftEnd)
    If aftRng.Font.Italic <> True Then
        IsInItalicBlock = False
        Exit Function
    End If

    ' Both sides are entirely italic — likely a block quote
    IsInItalicBlock = True
    On Error GoTo 0
End Function

' ────────────────────────────────────────────────────────────
'  Core: scan forward from a trigger word to determine the
'  extent of the legislative reference, then de-italicise it.
'  Returns True if a change was made.
' ────────────────────────────────────────────────────────────
Private Function DeItaliciseSpan(doc As Document, _
                                  rng As Range) As Boolean
    DeItaliciseSpan = False

    Dim spanStart As Long
    spanStart = rng.Start

    Dim paraEnd As Long
    paraEnd = rng.Paragraphs(1).Range.End

    Dim scanPos As Long
    scanPos = rng.End
    Dim spanEnd As Long
    spanEnd = rng.End

    Dim wordRng As Range
    Dim wordText As String
    Dim trimmed As String
    Dim ch As String
    Dim wordStart As Long
    Dim wordEnd As Long

    Do While scanPos < paraEnd
        ' Skip whitespace
        Do While scanPos < paraEnd
            Set wordRng = doc.Range(scanPos, scanPos + 1)
            ch = wordRng.Text
            If ch <> " " And ch <> Chr$(160) Then Exit Do
            scanPos = scanPos + 1
        Loop
        If scanPos >= paraEnd Then Exit Do

        ' Gather a word (non-space run)
        wordStart = scanPos
        Do While scanPos < paraEnd
            Set wordRng = doc.Range(scanPos, scanPos + 1)
            ch = wordRng.Text
            If ch = " " Or ch = Chr$(160) Or ch = vbCr _
               Or ch = Chr$(13) Or ch = Chr$(7) Then Exit Do
            scanPos = scanPos + 1
        Loop
        wordEnd = scanPos

        Set wordRng = doc.Range(wordStart, wordEnd)
        wordText = wordRng.Text
        trimmed = Trim$(wordText)
        If Len(trimmed) = 0 Then Exit Do

        ' Stop if this word is no longer italic
        If Not IsItalic(wordRng) Then Exit Do

        ' Stop before a word with more than 3 alpha characters
        If AlphaCount(trimmed) > 3 Then Exit Do

        ' Include this word in the span
        spanEnd = wordEnd
    Loop

    ' ── Block-quote guard ────────────────────────────────────
    If IsInItalicBlock(doc, spanStart, spanEnd) Then
        Exit Function
    End If

    ' ── Apply formatting change ──────────────────────────────
    If spanEnd > spanStart Then
        Dim target As Range
        Set target = doc.Range(spanStart, spanEnd)
        If target.Font.Italic = True Or _
           target.Font.Italic = wdUndefined Then
            target.Font.Italic = False
            DeItaliciseSpan = True
        End If
    End If
End Function

' ────────────────────────────────────────────────────────────
'  Core: find and de-italicise a specific phrase wherever it
'  appears in italics (not inside a block quote).
'  Returns the number of changes made.
' ────────────────────────────────────────────────────────────
Private Function DeItalicisePhrase(doc As Document, _
                                    phrase As String) As Long
    Dim hits As Long
    hits = 0

    Dim rng As Range
    Set rng = doc.Content

    With rng.Find
        .ClearFormatting
        .Text = phrase
        .Font.Italic = True
        .MatchCase = False
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWholeWord = False
        .MatchWildcards = False

        Do While .Execute
            ' Block-quote guard
            If Not IsInItalicBlock(doc, rng.Start, rng.End) Then
                rng.Font.Italic = False
                hits = hits + 1
            End If

            ' Move past this match
            rng.Start = rng.Start + Len(phrase)
            rng.End = doc.Content.End
        Loop
    End With

    DeItalicisePhrase = hits
End Function

' ════════════════════════════════════════════════════════════
'  MAIN ENTRY POINT
' ════════════════════════════════════════════════════════════
Public Sub DeItaliciseSectionReferences()
    Dim doc As Document
    Set doc = ActiveDocument

    ' ── Remember and enable tracked changes ──────────────────
    Dim origTrack As Boolean
    origTrack = doc.TrackRevisions
    doc.TrackRevisions = True

    Dim hitCount As Long
    hitCount = 0

    ' ── Trigger words (case-insensitive search) ──────────────
    Dim triggers As Variant
    triggers = Array("Section", "Regulation", "Article", "Paragraph")

    Dim t As Long
    For t = LBound(triggers) To UBound(triggers)
        Dim rng As Range
        Set rng = doc.Content

        With rng.Find
            .ClearFormatting
            .Text = CStr(triggers(t))
            .Font.Italic = True
            .MatchCase = False
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .MatchWholeWord = False
            .MatchWildcards = False

            Do While .Execute
                Dim spanEnd As Long
                spanEnd = rng.End

                If DeItaliciseSpan(doc, rng) Then
                    hitCount = hitCount + 1
                End If

                ' Move past this match
                rng.Start = rng.End
                rng.End = doc.Content.End
            Loop
        End With
    Next t

    ' ── Specific statutory phrases ───────────────────────────
    Dim phrases As Variant
    phrases = Array( _
        "Bank of Uganda Act, 1966", _
        "Capital Adequacy Regulations", _
        "FI (Amendment) Act", _
        "Liquidity Regulations", _
        "Bank of Uganda Act")
    ' Note: longest/most specific variants first so that
    ' "Bank of Uganda Act, 1966" is matched before
    ' "Bank of Uganda Act".

    Dim p As Long
    For p = LBound(phrases) To UBound(phrases)
        hitCount = hitCount + DeItalicisePhrase(doc, CStr(phrases(p)))
    Next p

    ' ── Restore original Track Changes setting ───────────────
    doc.TrackRevisions = origTrack

    MsgBox "Done. De-italicised " & hitCount & " item(s).", _
           vbInformation, "De-italicise Section References"
End Sub
