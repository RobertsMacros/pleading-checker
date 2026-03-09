Attribute VB_Name = "DeItaliciseSections"
' ============================================================
' DeItaliciseSections.bas
' Standalone macro: removes italic formatting from section
' references (e.g. "Section 18(2)", "Section 5 of the Act")
' using tracked changes.
'
' The span runs from "Section" forward through all words with
' three or fewer alphabetic characters (numbers, parentheticals,
' "of", "the", "Act" etc.), stopping before the first word
' with more than three alphabetic characters.
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
        ' Mixed formatting — check first character
        Dim cr As Range
        Set cr = rng.Document.Range(rng.Start, rng.Start + 1)
        IsItalic = (cr.Font.Italic = True)
    Else
        IsItalic = False
    End If
    On Error GoTo 0
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

    ' ── Search for italic "Section" ──────────────────────────
    Dim rng As Range
    Set rng = doc.Content

    With rng.Find
        .ClearFormatting
        .Text = "Section"
        .Font.Italic = True
        .MatchCase = True
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchWholeWord = False
        .MatchWildcards = False

        Do While .Execute
            ' rng now covers the word "Section"
            Dim spanStart As Long
            spanStart = rng.Start

            ' ── Scan forward to determine extent of reference ─
            Dim paraEnd As Long
            paraEnd = rng.Paragraphs(1).Range.End

            Dim scanPos As Long
            scanPos = rng.End          ' just past "Section"
            Dim spanEnd As Long
            spanEnd = rng.End          ' will grow as we include tokens

            ' Walk through subsequent words
            Dim wordRng As Range
            Dim wordText As String
            Dim trimmed As String

            Do While scanPos < paraEnd
                ' Skip whitespace / spaces
                Dim ch As String
                ch = ""
                Do While scanPos < paraEnd
                    Set wordRng = doc.Range(scanPos, scanPos + 1)
                    ch = wordRng.Text
                    If ch <> " " And ch <> Chr$(160) Then Exit Do
                    scanPos = scanPos + 1
                Loop
                If scanPos >= paraEnd Then Exit Do

                ' Gather a word (non-space run)
                Dim wordStart As Long
                wordStart = scanPos
                Do While scanPos < paraEnd
                    Set wordRng = doc.Range(scanPos, scanPos + 1)
                    ch = wordRng.Text
                    If ch = " " Or ch = Chr$(160) Or ch = vbCr _
                       Or ch = Chr$(13) Or ch = Chr$(7) Then Exit Do
                    scanPos = scanPos + 1
                Loop
                Dim wordEnd As Long
                wordEnd = scanPos

                ' Get the word text
                Set wordRng = doc.Range(wordStart, wordEnd)
                wordText = wordRng.Text
                trimmed = Trim$(wordText)
                If Len(trimmed) = 0 Then Exit Do

                ' Check whether this word is still italic
                If Not IsItalic(wordRng) Then Exit Do

                ' Count alpha characters
                If AlphaCount(trimmed) > 3 Then
                    ' This word is beyond the section reference — stop
                    Exit Do
                End If

                ' Include this word in the span
                spanEnd = wordEnd
            Loop

            ' ── Apply formatting change ──────────────────────────
            If spanEnd > spanStart Then
                Dim target As Range
                Set target = doc.Range(spanStart, spanEnd)

                ' Only act if at least part of it is italic
                If target.Font.Italic = True Or _
                   target.Font.Italic = wdUndefined Then
                    target.Font.Italic = False
                    hitCount = hitCount + 1
                End If
            End If

            ' ── Move past this match to continue searching ───────
            rng.Start = spanEnd
            rng.End = doc.Content.End
        Loop
    End With

    ' ── Restore original Track Changes setting ───────────────
    doc.TrackRevisions = origTrack

    MsgBox "Done. De-italicised " & hitCount & " section reference(s).", _
           vbInformation, "De-italicise Section References"
End Sub
