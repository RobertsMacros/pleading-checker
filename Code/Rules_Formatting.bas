Attribute VB_Name = "Rules_Formatting"
' ============================================================
' Rules_Formatting.bas
' Public helper: IsBlockQuotePara
'
' IsBlockQuotePara is a public helper used by other modules.
' It requires STRONG indicators beyond mere indentation:
'   - Quote-related style name (definitive)
'   - Indentation + quotation-mark wrapping
'   - Indentation + entirely italic text
' Indentation + smaller font alone is NOT sufficient.
'
' Previous rules (paragraph break consistency, font consistency)
' have been retired as part of the MVP pruning pass.
' ============================================================
Option Explicit

' ------------------------------------------------------------
'  PUBLIC: Detect block quote / indented extract paragraphs.
'
'  STRICT RULE: Indentation alone is NEVER enough.
'  Smaller font + indentation alone is NEVER enough.
'  A block quote must have at least one of:
'    1. A block-quote style (name contains "quote"/"block"/"extract")
'    2. Enclosing quotation marks AND indentation
'    3. Entirely italic text AND indentation
'  Lists, numbered paragraphs, and bullet items are explicitly excluded.
' ------------------------------------------------------------
Public Function IsBlockQuotePara(para As Paragraph) As Boolean
    IsBlockQuotePara = False
    On Error Resume Next

    ' ==========================================================
    '  CHECK 0: Exclude list paragraphs (numbered, bulleted, etc.)
    '  Lists must NEVER be treated as block quotes.
    ' ==========================================================
    Dim listLvl As Long
    listLvl = 0
    listLvl = para.Range.ListFormat.ListLevelNumber
    If Err.Number <> 0 Then listLvl = 0: Err.Clear
    ' ListLevelNumber > 0 means this paragraph is in a list
    If listLvl > 0 Then
        On Error GoTo 0
        Exit Function
    End If

    ' Also check for list-like text patterns (manual numbering)
    Dim pTextRaw As String
    pTextRaw = ""
    pTextRaw = para.Range.Text
    If Err.Number <> 0 Then pTextRaw = "": Err.Clear
    On Error GoTo 0
    Dim pTextTrimmed As String
    pTextTrimmed = Replace(Replace(Replace(pTextRaw, vbCr, ""), vbTab, ""), ChrW(160), " ")
    pTextTrimmed = Trim$(pTextTrimmed)

    ' Check for bullet-like or number-list-like starts
    If Len(pTextTrimmed) > 1 Then
        Dim firstTwo As String
        firstTwo = Left$(pTextTrimmed, 2)
        ' Bullet characters: bullet, en-dash, em-dash, hyphen
        If Left$(pTextTrimmed, 1) = ChrW(8226) Or _
           Left$(pTextTrimmed, 1) = ChrW(8211) & " " Or _
           firstTwo = "- " Or firstTwo = "* " Then
            On Error GoTo 0
            Exit Function
        End If
        ' Numbered list pattern: "(a)", "(i)", "(1)", "1.", "a.", "i."
        If pTextTrimmed Like "(#)*" Or pTextTrimmed Like "(##)*" Or _
           pTextTrimmed Like "([a-z])*" Or pTextTrimmed Like "([ivx])*" Or _
           pTextTrimmed Like "#.*" Or pTextTrimmed Like "##.*" Or _
           pTextTrimmed Like "[a-z].*" Then
            On Error GoTo 0
            Exit Function
        End If
    End If

    ' Also check ListFormat.ListString for auto-numbered lists
    Dim listStr As String
    listStr = ""
    On Error Resume Next
    listStr = para.Range.ListFormat.ListString
    If Err.Number <> 0 Then listStr = "": Err.Clear
    If Len(listStr) > 0 Then
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  CHECK 1: Style name for quote/block/extract keywords
    '  (Definitive indicator - no other checks needed)
    ' ==========================================================
    Dim sn As String
    sn = LCase(para.Style.NameLocal)
    If Err.Number <> 0 Then sn = "": Err.Clear
    If InStr(sn, "quote") > 0 Or InStr(sn, "block") > 0 Or _
       InStr(sn, "extract") > 0 Then
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  INDENTATION CHECK
    '  All remaining indicators require indentation.
    ' ==========================================================
    Dim leftInd As Single
    leftInd = para.Format.LeftIndent
    If Err.Number <> 0 Then leftInd = 0: Err.Clear
    On Error GoTo 0

    ' No indentation = not a block quote (style check already done above)
    If leftInd <= 18 Then
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  CHECK 2: Indentation + quotation marks wrapping
    '  Starts or ends with a quotation mark character.
    ' ==========================================================
    Dim startsWithQuote As Boolean
    Dim endsWithQuote As Boolean
    startsWithQuote = False
    endsWithQuote = False
    If Len(pTextTrimmed) > 1 Then
        Dim fcChar As String
        Dim lcChar As String
        fcChar = Left$(pTextTrimmed, 1)
        lcChar = Right$(pTextTrimmed, 1)
        startsWithQuote = (fcChar = Chr(34) Or fcChar = ChrW(8220) Or fcChar = ChrW(8216))
        endsWithQuote = (lcChar = Chr(34) Or lcChar = ChrW(8221) Or lcChar = ChrW(8217))
    End If

    ' Block quote if indented AND wrapped in quotation marks
    If startsWithQuote Or endsWithQuote Then
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  CHECK 3: Indentation + entirely italic
    '  wdTrue (-1) means ALL text in the range is italic.
    ' ==========================================================
    Dim italVal As Long
    On Error Resume Next
    italVal = para.Range.Font.Italic
    If Err.Number <> 0 Then italVal = 0: Err.Clear
    If italVal = -1 Then  ' wdTrue = -1 means ALL italic
        IsBlockQuotePara = True
        On Error GoTo 0
        Exit Function
    End If

    ' ==========================================================
    '  DEFAULT: Indented but no strong indicator = NOT a block quote.
    '  Smaller font + indentation alone is deliberately NOT enough.
    '  This prevents indented lists, definitions, and body text
    '  from being misclassified.
    ' ==========================================================

    On Error GoTo 0
End Function
