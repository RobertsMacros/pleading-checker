Attribute VB_Name = "WordChecker"
' ============================================================
' WordChecker.bas
' Legal document formatting and consistency checker for Word
' Covers: case names, spacing, i.e./e.g., edition abbreviations
' (footnotes only), ellipses, pinpoints, cross-references,
' citations, square brackets, tables, capitalisation, dashes,
' and legal blob placeholders.
'
' Installation:
'   1. Open the VBA Editor (Alt+F11)
'   2. File > Import File > select WordChecker.bas
'   3. File > Import File > select frmWordChecker.frm
'   4. Run the macro "WordChecker" (or assign to a ribbon button)
' ============================================================
Option Explicit

' ── Entry point ─────────────────────────────────────────────
Public Sub WordChecker()
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first.", vbExclamation, "Word Checker"
        Exit Sub
    End If
    frmWordChecker.Show
End Sub

' ════════════════════════════════════════════════════════════
'  CASE NAMES
' ════════════════════════════════════════════════════════════
' style: "underline", "italic", "both"
' vDot:  True = use "v.", False = use "v"
Public Function CheckCaseNames(doc As Document, _
                               style As String, _
                               vDot As Boolean) As String
    Dim rng As Range
    Dim fnd As Find
    Dim count As Long
    Dim targetV As String
    targetV = IIf(vDot, "v.", "v")

    ' Two passes: with dot and without dot
    Dim patterns(1) As String
    patterns(0) = "[A-Z][A-Za-z ,''\-]@ v\. [A-Z][A-Za-z ,''\-]@"
    patterns(1) = "[A-Z][A-Za-z ,''\-]@ v [A-Z][A-Za-z ,''\-]@"

    Dim p As Integer
    For p = 0 To 1
        Set rng = doc.Content
        Set fnd = rng.Find
        With fnd
            .ClearFormatting
            .Text = patterns(p)
            .MatchWildcards = True
            .MatchCase = True
            .Wrap = wdFindStop
            .Forward = True
        End With

        Do While fnd.Execute
            ' Apply chosen formatting
            Select Case LCase(style)
                Case "underline"
                    rng.Font.Underline = wdUnderlineSingle
                    rng.Font.Italic = False
                Case "italic"
                    rng.Font.Italic = True
                    rng.Font.Underline = wdUnderlineNone
                Case "both"
                    rng.Font.Underline = wdUnderlineSingle
                    rng.Font.Italic = True
            End Select

            ' Normalise "v" vs "v." within the match
            Dim innerRng As Range
            Set innerRng = rng.Duplicate
            With innerRng.Find
                .ClearFormatting
                .MatchWildcards = False
                .MatchWholeWord = True
                .MatchCase = True
                .Wrap = wdFindStop
                If vDot Then
                    .Text = " v "
                    .Replacement.Text = " v. "
                Else
                    .Text = " v. "
                    .Replacement.Text = " v "
                End If
                .Execute Replace:=wdReplaceAll
            End With

            count = count + 1
            ' Collapse to end to continue search
            rng.Collapse wdCollapseEnd
        Loop
    Next p

    CheckCaseNames = "Case names: " & count & " instance(s) formatted."
End Function

' ════════════════════════════════════════════════════════════
'  SPACING AFTER FULL STOP
' ════════════════════════════════════════════════════════════
' doubleSpace: True = enforce two spaces, False = enforce one
Public Function CheckSpacing(doc As Document, doubleSpace As Boolean) As String
    Dim fnd As Find

    If doubleSpace Then
        ' Step 1: collapse any existing multi-spaces down to single.
        ' Loop until stable so triple/quad+ spaces are handled too.
        Set fnd = doc.Content.Find
        With fnd
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ".  "       ' period + two spaces
            .Replacement.Text = ". "
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
        End With
        Do While fnd.Execute(Replace:=wdReplaceAll)
        Loop

        ' Step 2: expand every single post-period space to double.
        Set fnd = doc.Content.Find
        With fnd
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ". "
            .Replacement.Text = ".  "
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
            .Execute Replace:=wdReplaceAll
        End With
        CheckSpacing = "Spacing: double space after full stop enforced."
    Else
        ' Collapse to single: repeat until no ".  " remains (handles triple+).
        Set fnd = doc.Content.Find
        With fnd
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ".  "
            .Replacement.Text = ". "
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
        End With
        Do While fnd.Execute(Replace:=wdReplaceAll)
        Loop
        CheckSpacing = "Spacing: single space after full stop enforced."
    End If
End Function

' ════════════════════════════════════════════════════════════
'  i.e. AND e.g. STANDARDISATION
' ════════════════════════════════════════════════════════════
' ieConvention / egConvention: the desired form, e.g. "i.e." or "ie"
' ieComma / egComma: whether a comma should follow
'
' Two-pass approach to avoid double-comma bug:
'   Pass 1 — strip ALL trailing commas from every variant, normalise to base form.
'             Because commas are non-word characters, MatchWholeWord on "i.e." also
'             matches "i.e.," — so we must clear commas before adding them.
'   Pass 2 — add comma to base form if required.
Public Function CheckIeEg(doc As Document, _
                           ieConvention As String, _
                           egConvention As String, _
                           ieComma As Boolean, _
                           egComma As Boolean) As String

    ' Strip any trailing comma the user may have typed in the convention box
    Dim ieBase As String: ieBase = Trim(ieConvention)
    Dim egBase As String: egBase = Trim(egConvention)
    If Right(ieBase, 1) = "," Then ieBase = Left(ieBase, Len(ieBase) - 1)
    If Right(egBase, 1) = "," Then egBase = Left(egBase, Len(egBase) - 1)

    ' Guard: empty convention box would delete every variant from the document
    If Len(ieBase) = 0 Or Len(egBase) = 0 Then
        CheckIeEg = "i.e./e.g.: skipped — convention text box is empty."
        Exit Function
    End If

    Dim ieVariants(3) As String
    ieVariants(0) = "i.e."
    ieVariants(1) = "i.e"
    ieVariants(2) = "ie."
    ieVariants(3) = "ie"

    Dim egVariants(3) As String
    egVariants(0) = "e.g."
    egVariants(1) = "e.g"
    egVariants(2) = "eg."
    egVariants(3) = "eg"

    Dim i As Integer

    ' ── i.e. ────────────────────────────────────────────────
    ' Pass 1a: replace every variant+comma → ieBase  (comma stripping)
    For i = 0 To 3
        ReplaceWholeWord doc, ieVariants(i) & ",", ieBase
    Next i
    ' Pass 1b: replace every bare variant → ieBase
    For i = 0 To 3
        If ieVariants(i) <> ieBase Then
            ReplaceWholeWord doc, ieVariants(i), ieBase
        End If
    Next i
    ' Pass 2: add comma if required (safe — no commas remain from pass 1)
    If ieComma Then ReplaceWholeWord doc, ieBase, ieBase & ","

    ' ── e.g. ────────────────────────────────────────────────
    ' Pass 1a
    For i = 0 To 3
        ReplaceWholeWord doc, egVariants(i) & ",", egBase
    Next i
    ' Pass 1b
    For i = 0 To 3
        If egVariants(i) <> egBase Then
            ReplaceWholeWord doc, egVariants(i), egBase
        End If
    Next i
    ' Pass 2
    If egComma Then ReplaceWholeWord doc, egBase, egBase & ","

    CheckIeEg = "i.e./e.g.: conventions applied throughout."
End Function

' Helper: whole-word find/replace (skips if texts are identical or findText is empty)
Private Sub ReplaceWholeWord(doc As Document, _
                              findText As String, _
                              replText As String)
    If findText = replText Then Exit Sub
    If Len(findText) = 0 Then Exit Sub   ' guard: empty find would corrupt the document
    Dim fnd As Find
    Set fnd = doc.Content.Find
    With fnd
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replText
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' ════════════════════════════════════════════════════════════
'  EDITION ABBREVIATION  (footnotes and endnotes only)
' ════════════════════════════════════════════════════════════
' edStyle: "ed", "edn", or "edition"
Public Function CheckEdition(doc As Document, edStyle As String) As String
    Dim variants(5) As String
    variants(0) = "edition"
    variants(1) = "edn."
    variants(2) = "edn"
    variants(3) = "ed."
    variants(4) = "ed"
    variants(5) = "Edition"

    Dim i As Integer
    Dim target As String: target = edStyle

    ' Process footnote ranges
    Dim fn As Footnote
    For Each fn In doc.Footnotes
        For i = 0 To 5
            If LCase(variants(i)) <> LCase(target) Then
                ReplaceInRange fn.Range, variants(i), target
            End If
        Next i
    Next fn

    ' Process endnote ranges
    Dim en As Endnote
    For Each en In doc.Endnotes
        For i = 0 To 5
            If LCase(variants(i)) <> LCase(target) Then
                ReplaceInRange en.Range, variants(i), target
            End If
        Next i
    Next en

    CheckEdition = "Edition: """ & target & """ convention applied to footnotes and endnotes."
End Function

' Helper: replace within a specific range object
' (Execute(Replace:=wdReplaceAll) returns True/False, not a count,
'  so we do not attempt to return a replacement count from this helper.)
Private Sub ReplaceInRange(rng As Range, _
                            findText As String, _
                            replText As String)
    If Len(findText) = 0 Then Exit Sub
    Dim fnd As Find
    Set fnd = rng.Find
    With fnd
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replText
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' ════════════════════════════════════════════════════════════
'  ELLIPSES
' ════════════════════════════════════════════════════════════
Public Function CheckEllipses(doc As Document) As String
    Dim fnd As Find

    ' 1. Remove space BEFORE ellipsis: " ..." → "..."
    Set fnd = doc.Content.Find
    With fnd
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = " ..."
        .Replacement.Text = "..."
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
        .Execute Replace:=wdReplaceAll
    End With

    ' 2. Ensure space AFTER ellipsis when followed by a letter/digit (not end of sentence)
    '    Pattern: "..." not followed by space, closing punctuation, or end
    '    Use wildcard: "...([! .?!,;:)\]])" → "... \1"
    Set fnd = doc.Content.Find
    With fnd
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "\.\.\.([! .,;:?!\)\]])"
        .Replacement.Text = "... \1"
        .MatchWildcards = True
        .Wrap = wdFindStop
        .Forward = True
        .Execute Replace:=wdReplaceAll
    End With

    CheckEllipses = "Ellipses: spacing normalised (space before removed; space after added where needed)."
End Function

' ════════════════════════════════════════════════════════════
'  PINPOINT CITATIONS  ("at N" → ", N")
' ════════════════════════════════════════════════════════════
Public Function CheckPinpoints(doc As Document) As String
    Dim fnd As Find
    Set fnd = doc.Content.Find
    With fnd
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = " at ([0-9])"
        .Replacement.Text = ", \1"
        .MatchWildcards = True
        .Wrap = wdFindStop
        .Forward = True
        .Execute Replace:=wdReplaceAll
    End With
    CheckPinpoints = "Pinpoints: "" at N"" replaced with "", N"" throughout."
End Function

' ════════════════════════════════════════════════════════════
'  CROSS-REFERENCES (supra / infra / ibid)
' ════════════════════════════════════════════════════════════
Public Function CheckCrossRefs(doc As Document, noSupraInfra As Boolean) As String
    If Not noSupraInfra Then
        CheckCrossRefs = "Cross-references: check skipped (option not enabled)."
        Exit Function
    End If

    Dim terms(2) As String
    terms(0) = "supra"
    terms(1) = "infra"
    terms(2) = "ibid"

    Dim rng As Range
    Dim count As Long
    Dim i As Integer

    For i = 0 To 2
        Set rng = doc.Content
        With rng.Find
            .ClearFormatting
            .Text = terms(i)
            .MatchWholeWord = True
            .MatchCase = False
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
        End With
        Do While rng.Find.Execute
            rng.HighlightColorIndex = wdYellow
            count = count + 1
            rng.Collapse wdCollapseEnd
        Loop
    Next i

    CheckCrossRefs = "Cross-references: " & count & " instance(s) of supra/infra/ibid highlighted for review."
End Function

' ════════════════════════════════════════════════════════════
'  CITATION & EXHIBIT CONSISTENCY
' ════════════════════════════════════════════════════════════
Public Function CheckCitations(doc As Document, citFormat As String) As String
    ' Check exhibit prefix conventions: C-, CL-, CLA-, R-, RL-, RLA-
    Dim prefixes(5) As String
    prefixes(0) = "C-"
    prefixes(1) = "CL-"
    prefixes(2) = "CLA-"
    prefixes(3) = "R-"
    prefixes(4) = "RL-"
    prefixes(5) = "RLA-"

    Dim rng As Range
    Dim count As Long
    Dim i As Integer

    ' Highlight any exhibit references that don't match a recognised prefix
    ' We look for patterns like "[Exhibit X-NNN]" or "Exhibit X-NNN"
    For i = 0 To 5
        Set rng = doc.Content
        With rng.Find
            .ClearFormatting
            .Text = prefixes(i) & "[0-9]@"
            .MatchWildcards = True
            .Wrap = wdFindStop
            .Forward = True
        End With
        Do While rng.Find.Execute
            count = count + 1
            rng.Collapse wdCollapseEnd
        Loop
    Next i

    CheckCitations = "Citations: " & count & " exhibit reference(s) found matching recognised prefixes."
End Function

' ════════════════════════════════════════════════════════════
'  SQUARE BRACKETS IN QUOTATIONS (must not be italic)
' ════════════════════════════════════════════════════════════
Public Function CheckBracketsInQuotes(doc As Document) As String
    Dim rng As Range
    Dim count As Long
    Dim chars() As String
    chars = Split("[ ]", " ")  ' open and close brackets

    Dim c As Integer
    For c = 0 To 1
        Set rng = doc.Content
        With rng.Find
            .ClearFormatting
            .Font.Italic = True
            .Text = chars(c)
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
        End With
        Do While rng.Find.Execute
            rng.Font.Italic = False
            count = count + 1
            rng.Collapse wdCollapseEnd
        Loop
    Next c

    CheckBracketsInQuotes = "Square brackets: " & count & " italic bracket(s) de-italicised."
End Function

' ════════════════════════════════════════════════════════════
'  TABLE SPACING (6pt above and below)
' ════════════════════════════════════════════════════════════
Public Function CheckTables(doc As Document) As String
    Dim tbl As Table
    Dim cel As Cell
    Dim para As Paragraph
    Dim flagCount As Long
    Dim fixCount As Long

    For Each tbl In doc.Tables
        For Each cel In tbl.Range.Cells
            For Each para In cel.Range.Paragraphs
                Dim needsFix As Boolean: needsFix = False
                If para.Format.SpaceBefore <> 6 Then needsFix = True
                If para.Format.SpaceAfter <> 6 Then needsFix = True
                If needsFix Then
                    para.Format.SpaceBefore = 6
                    para.Format.SpaceAfter = 6
                    fixCount = fixCount + 1
                End If
            Next para
        Next cel
    Next tbl

    If fixCount = 0 Then
        CheckTables = "Tables: all paragraph spacing already correct (6pt above/below)."
    Else
        CheckTables = "Tables: " & fixCount & " paragraph(s) updated to 6pt above/below."
    End If
End Function

' ════════════════════════════════════════════════════════════
'  CAPITALISATION (e.g. clause / Clause)
' Opens a sub-dialog if both forms exist so the user can choose.
' ════════════════════════════════════════════════════════════
Public Function CheckCapitalisation(doc As Document) As String
    ' Term pairs: (lowercase, Capitalised)
    Dim terms(3, 1) As String
    terms(0, 0) = "clause":   terms(0, 1) = "Clause"
    terms(1, 0) = "schedule": terms(1, 1) = "Schedule"
    terms(2, 0) = "exhibit":  terms(2, 1) = "Exhibit"
    terms(3, 0) = "paragraph": terms(3, 1) = "Paragraph"

    Dim results As String
    Dim i As Integer

    For i = 0 To 3
        Dim lcWord As String: lcWord = terms(i, 0)
        Dim ucWord As String: ucWord = terms(i, 1)
        Dim lcCount As Long: lcCount = CountWhole(doc, lcWord)
        Dim ucCount As Long: ucCount = CountWhole(doc, ucWord)

        If lcCount > 0 And ucCount > 0 Then
            ' Both forms exist — ask user
            Dim msg As String
            msg = "Found both """ & lcWord & """ (" & lcCount & "x) and """ & _
                  ucWord & """ (" & ucCount & "x)." & vbCrLf & vbCrLf & _
                  "Which form do you want to use throughout?"
            Dim answer As VbMsgBoxResult
            answer = MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton1, _
                            "Capitalisation: " & ucWord & " / " & lcWord & _
                            " — click Yes for """ & ucWord & """, No for """ & lcWord & """")
            If answer = vbYes Then
                ' Standardise to Capitalised form
                ReplaceWholeWord doc, lcWord, ucWord
                results = results & ucWord & ": standardised to capitalised. "
            Else
                ' Standardise to lowercase form
                ReplaceWholeWord doc, ucWord, lcWord
                results = results & lcWord & ": standardised to lowercase. "
            End If
        ElseIf lcCount > 0 And ucCount = 0 Then
            results = results & lcWord & ": consistent (lowercase). "
        ElseIf ucCount > 0 And lcCount = 0 Then
            results = results & ucWord & ": consistent (capitalised). "
        End If
    Next i

    If results = "" Then results = "no issues found."
    CheckCapitalisation = "Capitalisation: " & results
End Function

' Helper: count whole-word occurrences (case-sensitive)
Private Function CountWhole(doc As Document, word As String) As Long
    Dim rng As Range
    Set rng = doc.Content
    Dim count As Long
    With rng.Find
        .ClearFormatting
        .Text = word
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
    End With
    Do While rng.Find.Execute
        count = count + 1
        rng.Collapse wdCollapseEnd
    Loop
    CountWhole = count
End Function

' ════════════════════════════════════════════════════════════
'  DASHES  (en dash / em dash / hyphen)
' ════════════════════════════════════════════════════════════
Public Function CheckDashes(doc As Document) As String
    Dim enDash As String:  enDash  = ChrW(8211)  ' –  U+2013
    Dim emDash As String:  emDash  = ChrW(8212)  ' —  U+2014

    ' 1. Numeric ranges: digit-hyphen-digit → digit–digit (en dash)
    '    Note: Execute(Replace:=wdReplaceAll) returns True/False, not a count,
    '    so we do not attempt to capture it as a number.
    Dim rng As Range
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "([0-9])\-([0-9])"
        .Replacement.Text = "\1" & enDash & "\2"
        .MatchWildcards = True
        .Wrap = wdFindStop
        .Forward = True
        .Execute Replace:=wdReplaceAll
    End With

    ' 2. Spaced hyphen used as parenthetical dash: " - " → " — " (em dash)
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = " - "
        .Replacement.Text = " " & emDash & " "
        .MatchWildcards = False
        .Wrap = wdFindStop
        .Forward = True
        .Execute Replace:=wdReplaceAll
    End With

    CheckDashes = "Dashes: numeric range hyphens converted to en dash (–); " & _
                  "spaced hyphens converted to em dash (—)."
End Function

' ════════════════════════════════════════════════════════════
'  LEGAL BLOBS  (unfilled placeholders: ·  •  ●  and variants)
' ════════════════════════════════════════════════════════════
Public Function CheckLegalBlobs(doc As Document) As String
    ' Three blob characters + bracketed variants
    Dim blobChars(2) As String
    blobChars(0) = ChrW(183)   ' · middle dot / thin blob   U+00B7
    blobChars(1) = ChrW(8226)  ' • bullet / medium blob     U+2022
    blobChars(2) = ChrW(9679)  ' ● large circle / fat blob  U+25CF

    Dim blobNames(2) As String
    blobNames(0) = "thin (·)"
    blobNames(1) = "medium (•)"
    blobNames(2) = "fat (●)"

    Dim counts(2) As Long
    Dim total As Long
    Dim i As Integer
    Dim rng As Range

    For i = 0 To 2
        ' Search for bare blob
        Set rng = doc.Content
        With rng.Find
            .ClearFormatting
            .Text = blobChars(i)
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
        End With
        Do While rng.Find.Execute
            rng.HighlightColorIndex = wdYellow
            counts(i) = counts(i) + 1
            total = total + 1
            rng.Collapse wdCollapseEnd
        Loop

        ' Search for bracket-wrapped blob: [·] [•] [●]
        Set rng = doc.Content
        With rng.Find
            .ClearFormatting
            .Text = "[" & blobChars(i) & "]"
            .MatchWildcards = False
            .Wrap = wdFindStop
            .Forward = True
        End With
        Do While rng.Find.Execute
            rng.HighlightColorIndex = wdYellow
            counts(i) = counts(i) + 1
            total = total + 1
            rng.Collapse wdCollapseEnd
        Loop
    Next i

    If total = 0 Then
        CheckLegalBlobs = "Blobs: none found — document is blob-free."
    Else
        Dim detail As String
        For i = 0 To 2
            If counts(i) > 0 Then
                detail = detail & counts(i) & "x " & blobNames(i) & "  "
            End If
        Next i
        CheckLegalBlobs = "*** BLOBS FOUND: " & total & " unfilled placeholder(s) highlighted. " & _
                          "(" & Trim(detail) & ") — DO NOT publish until resolved. ***"
    End If
End Function

' ════════════════════════════════════════════════════════════
'  TRACKED CHANGES TOGGLE
' ════════════════════════════════════════════════════════════
Public Sub ToggleTrackedChanges(doc As Document)
    doc.TrackRevisions = Not doc.TrackRevisions
End Sub

Public Function TrackChangesCaption(doc As Document) As String
    If doc.TrackRevisions Then
        TrackChangesCaption = "Track Changes: ON"
    Else
        TrackChangesCaption = "Track Changes: OFF"
    End If
End Function

' ════════════════════════════════════════════════════════════
'  DOCUMENT ACTIONS
' ════════════════════════════════════════════════════════════
Public Sub UpdateTOC(doc As Document)
    Dim toc As TableOfContents
    For Each toc In doc.TablesOfContents
        toc.Update
    Next toc
End Sub

Public Sub UpdateCrossRefs(doc As Document)
    doc.Fields.Update
End Sub

' ════════════════════════════════════════════════════════════
'  DRAFT EMAIL TO DOCUMENT SERVICES
' ════════════════════════════════════════════════════════════
Public Sub DraftDSEmail(doc As Document)
    Const DS_ADDRESS As String = "globaldocumentspecialists@freshfields.com"
    Const DS_DISPLAY As String = "Global Document Specialists"

    Dim docName As String
    docName = doc.Name

    Dim subjectLine As String
    subjectLine = docName & " — please format"

    Dim bodyText As String
    bodyText = "Hi DS" & vbCrLf & vbCrLf & _
               "I would be very grateful for a document tidy up of the attached by tomorrow morning at 9am." & _
               vbCrLf & vbCrLf & _
               "Many thanks"

    ' Try Outlook first
    On Error GoTo NoOutlook
    Dim olApp As Object
    Set olApp = CreateObject("Outlook.Application")

    Dim mail As Object
    Set mail = olApp.CreateItem(0)  ' olMailItem = 0
    With mail
        .To = DS_DISPLAY & " <" & DS_ADDRESS & ">"
        .Subject = subjectLine
        .Body = bodyText
        On Error Resume Next
        .Attachments.Add doc.FullName
        On Error GoTo 0
        .Display  ' opens draft for review rather than sending
    End With
    Exit Sub

NoOutlook:
    ' Fallback: copy to clipboard
    Dim fullText As String
    fullText = "To: " & DS_DISPLAY & " <" & DS_ADDRESS & ">" & vbCrLf & _
               "Subject: " & subjectLine & vbCrLf & vbCrLf & bodyText

    Dim dataObj As Object
    On Error Resume Next
    Set dataObj = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    If Err.Number = 0 Then
        dataObj.SetText fullText
        dataObj.PutInClipboard
        MsgBox "Outlook is not available." & vbCrLf & vbCrLf & _
               "The email text has been copied to your clipboard. " & _
               "Please paste it into your email client.", _
               vbInformation, "Draft DS Email"
    Else
        MsgBox "Outlook is not available. Please send the following manually:" & _
               vbCrLf & vbCrLf & fullText, vbInformation, "Draft DS Email"
    End If
    On Error GoTo 0
End Sub

' ════════════════════════════════════════════════════════════
'  RUN ALL CHECKS
' Returns a combined results string
' ════════════════════════════════════════════════════════════
Public Function RunAllChecks(doc As Document, _
                              caseStyle As String, _
                              vDot As Boolean, _
                              doubleSpace As Boolean, _
                              ieConvention As String, _
                              egConvention As String, _
                              ieComma As Boolean, _
                              egComma As Boolean, _
                              edStyle As String, _
                              noSupraInfra As Boolean, _
                              citFormat As String) As String
    Dim results As String
    Dim sep As String: sep = vbCrLf
    Dim r As String

    ' Each check is individually wrapped so that one failure never silences the rest.
    ' Execute(Replace:=wdReplaceAll) returns True/False, but we guard against any
    ' unexpected runtime errors (locked document, corrupt content, etc.).

    On Error Resume Next: Err.Clear
    r = CheckLegalBlobs(doc)    ' Blobs first — critical
    If Err.Number <> 0 Then r = "Blobs: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = r & sep

    On Error Resume Next: Err.Clear
    r = CheckCaseNames(doc, caseStyle, vDot)
    If Err.Number <> 0 Then r = "Case names: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    On Error Resume Next: Err.Clear
    r = CheckSpacing(doc, doubleSpace)
    If Err.Number <> 0 Then r = "Spacing: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    On Error Resume Next: Err.Clear
    r = CheckIeEg(doc, ieConvention, egConvention, ieComma, egComma)
    If Err.Number <> 0 Then r = "i.e./e.g.: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    On Error Resume Next: Err.Clear
    r = CheckEdition(doc, edStyle)
    If Err.Number <> 0 Then r = "Edition: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    On Error Resume Next: Err.Clear
    r = CheckEllipses(doc)
    If Err.Number <> 0 Then r = "Ellipses: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    On Error Resume Next: Err.Clear
    r = CheckPinpoints(doc)
    If Err.Number <> 0 Then r = "Pinpoints: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    On Error Resume Next: Err.Clear
    r = CheckCrossRefs(doc, noSupraInfra)
    If Err.Number <> 0 Then r = "Cross-references: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    On Error Resume Next: Err.Clear
    r = CheckCitations(doc, citFormat)
    If Err.Number <> 0 Then r = "Citations: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    On Error Resume Next: Err.Clear
    r = CheckBracketsInQuotes(doc)
    If Err.Number <> 0 Then r = "Brackets: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    On Error Resume Next: Err.Clear
    r = CheckTables(doc)
    If Err.Number <> 0 Then r = "Tables: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    On Error Resume Next: Err.Clear
    r = CheckCapitalisation(doc)
    If Err.Number <> 0 Then r = "Capitalisation: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    On Error Resume Next: Err.Clear
    r = CheckDashes(doc)
    If Err.Number <> 0 Then r = "Dashes: check failed (" & Err.Description & ")."
    On Error GoTo 0
    results = results & r & sep

    RunAllChecks = results
End Function
