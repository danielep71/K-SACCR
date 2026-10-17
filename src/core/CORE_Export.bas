Attribute VB_Name = "CORE_Export"
'==============================================================================
' MODULE: CORE_Export
'------------------------------------------------------------------------------
' PURPOSE
'   Format the fields and lines of the evidence report of a run as CSV text
'   and encode the report as UTF-8 bytes (#45), the same on every
'   computer: fields separated by ";", numbers with a decimal point and no
'   thousands separator, logical values as TRUE and FALSE, error values by
'   their Excel names, text quoted only when it holds a separator, a quote
'   or a line break, text that Excel would read as a formula prefixed with
'   an apostrophe, and lines ended by CR LF. Each line starts with its
'   section: RUN, PARAMS, SF_TABLE, FX_TABLE, RESULTS, RECONCILIATION,
'   TRADE_RECONCILIATION, CHECKS or LIMITATIONS.
'
' PUBLIC SURFACE
'   None outside this VBA project. NumberText, CsvField, CsvLine,
'   TableLines, LimitationLines and Utf8Bytes are Public for M_Main,
'   CORE_Engine, CORE_Reconcile and the tests; Option Private Module keeps
'   them off the supported external surface.
'
' DEPENDENCIES
'   None outside VBA.
'
' STATE OWNERSHIP
'   None. Every procedure works on its arguments only and touches no
'   workbook or file: M_Main reads the sheets and writes the file.
'
' ERROR POLICY
'   Unexpected errors are raised to the caller.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' UPDATED
'   2026-10-17
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the export inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
        Private Const SEP        As String = ";"       'Field separator
        Private Const QUOTE_CHR  As String = """"      'Text qualifier


'
'------------------------------------------------------------------------------
'
'                                   FIELDS
'
'------------------------------------------------------------------------------
'

Public Function NumberText( _
    ByVal Number As Double) _
    As String
'
'==============================================================================
'                                  NumberText
'------------------------------------------------------------------------------
' PURPOSE
'   Write a number with a decimal point, no thousands separator and up to
'   15 significant digits, whatever the locale: 1234.5, 0.01, -0.5, 1E+20.
'
' METHOD
'   Str$ always uses a decimal point; it leaves a leading space for a
'   positive number and no zero before the point, which is added here.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Formatted   As String    'The number as Str$ writes it

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        Formatted = Trim$(Str$(Number))
        If Left$(Formatted, 1) = "." Then
            Formatted = "0" & Formatted
        ElseIf Left$(Formatted, 2) = "-." Then
            Formatted = "-0" & Mid$(Formatted, 2)
        End If
        NumberText = Formatted

End Function


Public Function CsvField( _
    ByVal Content As Variant) _
    As String
'
'==============================================================================
'                                   CsvField
'------------------------------------------------------------------------------
' PURPOSE
'   Write one cell value (Value2) as a CSV field.
'
' RETURNS
'   "" for an empty cell; NumberText for a number, and for a date its
'   serial number; TRUE or FALSE; the Excel name of an error value (#N/A,
'   #VALUE!, ...); text as it is, in double quotes with each quote doubled
'   when it holds ";", a quote or a line break. Text that Excel would read
'   as a formula when the file is opened (starting with =, +, -, @, a tab
'   or a carriage return) is written with a leading apostrophe; numbers
'   are not text and keep their sign.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Field   As String    'Text of the value

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        Select Case VarType(Content)
            Case vbEmpty, vbNull
                CsvField = ""
            Case vbBoolean
                CsvField = IIf(Content, "TRUE", "FALSE")
            Case vbError
                CsvField = ErrorName(Content)
            Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte, vbDate
                CsvField = NumberText(CDbl(Content))
            Case Else
                Field = CStr(Content)
                If Len(Field) > 0 Then
                    If InStr("=+-@" & vbTab & vbCr, Left$(Field, 1)) > 0 Then
                        Field = "'" & Field
                    End If
                End If
                If InStr(Field, SEP) > 0 Or InStr(Field, QUOTE_CHR) > 0 Or InStr(Field, vbCr) > 0 Or InStr(Field, vbLf) > 0 Then
                    Field = QUOTE_CHR & Replace(Field, QUOTE_CHR, QUOTE_CHR & QUOTE_CHR) & QUOTE_CHR
                End If
                CsvField = Field
        End Select

End Function


Public Function CsvLine( _
    ByVal Section As String, _
    ByVal Fields As Variant) _
    As String
'
'==============================================================================
'                                   CsvLine
'------------------------------------------------------------------------------
' PURPOSE
'   Write one line: the section, then each field, ended by CR LF.
'
' INPUTS
'   Section: RUN, PARAMS, ...
'   Fields: a 1-D array of values, with any lower bound.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parts()   As String    'Section and fields
    Dim i         As Long      'Field

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        ReDim Parts(0 To UBound(Fields) - LBound(Fields) + 1)
        Parts(0) = CsvField(Section)
        For i = LBound(Fields) To UBound(Fields)
            Parts(i - LBound(Fields) + 1) = CsvField(Fields(i))
        Next i
        CsvLine = Join(Parts, SEP) & vbCrLf

End Function


Public Function TableLines( _
    ByVal Section As String, _
    ByVal Block As Variant) _
    As String
'
'==============================================================================
'                                  TableLines
'------------------------------------------------------------------------------
' PURPOSE
'   Write every row of a 2-D block as one line of a section, in order;
'   the block's first row is normally its headers.
'
' INPUTS
'   Section: the section of every line.
'   Block: a 2-D array, as Range.Value2 of a range of several cells gives
'   one; its bounds may start at any number.
'
' RETURNS
'   The lines, or "" when Block is not an array (Value2 of a single cell).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lines()       As String     'One line per row
    Dim Fields()      As Variant    'Fields of one row
    Dim r             As Long       'Row
    Dim c             As Long       'Column
    Dim ColumnCount   As Long       'Columns

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        If Not IsArray(Block) Then
            Exit Function
        End If
        ColumnCount = UBound(Block, 2) - LBound(Block, 2) + 1
        ReDim Lines(LBound(Block, 1) To UBound(Block, 1))
        ReDim Fields(1 To ColumnCount)
        For r = LBound(Block, 1) To UBound(Block, 1)
            For c = 1 To ColumnCount
                Fields(c) = Block(r, LBound(Block, 2) + c - 1)
            Next c
            Lines(r) = CsvLine(Section, Fields)
        Next r
        TableLines = Join(Lines, "")

End Function


Public Function LimitationLines() As String
'
'==============================================================================
'                               LimitationLines
'------------------------------------------------------------------------------
' PURPOSE
'   The fixed LIMITATIONS section of every evidence report: what the
'   report does and does not show.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Notes   As Variant    'One limitation per line
    Dim i       As Long       'Limitation
    Dim Lines   As String     'The lines

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        Notes = Array( _
            "The report describes one run of this workbook on the inputs fingerprinted in RUN; it is not a " & _
                "regulatory certification.", _
            "A comparison with a bank's figures is not independent validation of either (#100).", _
            "Expected values classed illustrative show a behaviour and validate nothing.", _
            "Totals cover the VALID netting sets of one regime only; regimes are never added together.", _
            "Netting sets that are INVALID, INCOMPLETE or NO TRADES have no EAD and are listed on RESULTS " & _
                "and CHECKS.", _
            "Bank amounts are scaled to units as stated in RUN; ratios and times are not scaled.", _
            "Lines holding a run or reconciliation time change at every run; every other line depends only " & _
                "on the inputs.")
        For i = LBound(Notes) To UBound(Notes)
            Lines = Lines & CsvLine("LIMITATIONS", Array(i + 1, Notes(i)))
        Next i
        LimitationLines = Lines

End Function


'
'------------------------------------------------------------------------------
'
'                                    BYTES
'
'------------------------------------------------------------------------------
'

Public Function Utf8Bytes( _
    ByVal Text As String) _
    As Byte()
'
'==============================================================================
'                                  Utf8Bytes
'------------------------------------------------------------------------------
' PURPOSE
'   Encode text as UTF-8, without a byte-order mark and without any
'   library, so that the file is the same on every computer.
'
' METHOD
'   Each UTF-16 code unit of a VBA string is read with AscW; a surrogate
'   pair becomes one 4-byte sequence, and a lone surrogate becomes U+FFFD.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Bytes()        As Byte    'Bytes written
    Dim ByteCount      As Long    'Bytes used
    Dim i              As Long    'Position in text
    Dim CodePoint      As Long    'Code point
    Dim LowSurrogate   As Long    'Low surrogate

'------------------------------------------------------------------------------
' ENCODE
'------------------------------------------------------------------------------
    'An empty text gives an empty array. Three bytes per UTF-16 code unit
    'are always enough: a surrogate pair, two units, needs four.
        If Len(Text) = 0 Then
            Utf8Bytes = Bytes
            Exit Function
        End If
        ReDim Bytes(0 To Len(Text) * 3 - 1)
        i = 1
        Do While i <= Len(Text)
            CodePoint = AscW(Mid$(Text, i, 1)) And &HFFFF&
            If CodePoint >= &HD800& And CodePoint <= &HDBFF& And i < Len(Text) Then
                LowSurrogate = AscW(Mid$(Text, i + 1, 1)) And &HFFFF&
                If LowSurrogate >= &HDC00& And LowSurrogate <= &HDFFF& Then
                    CodePoint = &H10000 + (CodePoint - &HD800&) * &H400& + (LowSurrogate - &HDC00&)
                    i = i + 1
                Else
                    CodePoint = &HFFFD&
                End If
            ElseIf CodePoint >= &HD800& And CodePoint <= &HDFFF& Then
                CodePoint = &HFFFD&
            End If
            If CodePoint < &H80& Then
                Bytes(ByteCount) = CodePoint
                ByteCount = ByteCount + 1
            ElseIf CodePoint < &H800& Then
                Bytes(ByteCount) = &HC0& Or (CodePoint \ &H40&)
                Bytes(ByteCount + 1) = &H80& Or (CodePoint And &H3F&)
                ByteCount = ByteCount + 2
            ElseIf CodePoint < &H10000 Then
                Bytes(ByteCount) = &HE0& Or (CodePoint \ &H1000&)
                Bytes(ByteCount + 1) = &H80& Or ((CodePoint \ &H40&) And &H3F&)
                Bytes(ByteCount + 2) = &H80& Or (CodePoint And &H3F&)
                ByteCount = ByteCount + 3
            Else
                Bytes(ByteCount) = &HF0& Or (CodePoint \ &H40000)
                Bytes(ByteCount + 1) = &H80& Or ((CodePoint \ &H1000&) And &H3F&)
                Bytes(ByteCount + 2) = &H80& Or ((CodePoint \ &H40&) And &H3F&)
                Bytes(ByteCount + 3) = &H80& Or (CodePoint And &H3F&)
                ByteCount = ByteCount + 4
            End If
            i = i + 1
        Loop
        ReDim Preserve Bytes(0 To ByteCount - 1)
        Utf8Bytes = Bytes

End Function


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function ErrorName( _
    ByVal ErrorValue As Variant) _
    As String
'
'==============================================================================
'                                  ErrorName
'------------------------------------------------------------------------------
' PURPOSE
'   Return the Excel name of an error value, the same in every language.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
    'CLng of an error value raises, so the value is compared with each
    'error in turn; one not listed is written "#Error nnnn", from CStr.
        If ErrorValue = CVErr(2000) Then
            ErrorName = "#NULL!"
        ElseIf ErrorValue = CVErr(2007) Then
            ErrorName = "#DIV/0!"
        ElseIf ErrorValue = CVErr(2015) Then
            ErrorName = "#VALUE!"
        ElseIf ErrorValue = CVErr(2023) Then
            ErrorName = "#REF!"
        ElseIf ErrorValue = CVErr(2029) Then
            ErrorName = "#NAME?"
        ElseIf ErrorValue = CVErr(2036) Then
            ErrorName = "#NUM!"
        ElseIf ErrorValue = CVErr(2042) Then
            ErrorName = "#N/A"
        ElseIf ErrorValue = CVErr(2043) Then
            ErrorName = "#GETTING_DATA"
        ElseIf ErrorValue = CVErr(2045) Then
            ErrorName = "#SPILL!"
        ElseIf ErrorValue = CVErr(2046) Then
            ErrorName = "#CONNECT!"
        ElseIf ErrorValue = CVErr(2047) Then
            ErrorName = "#BLOCKED!"
        ElseIf ErrorValue = CVErr(2048) Then
            ErrorName = "#UNKNOWN!"
        ElseIf ErrorValue = CVErr(2049) Then
            ErrorName = "#FIELD!"
        ElseIf ErrorValue = CVErr(2050) Then
            ErrorName = "#CALC!"
        Else
            ErrorName = "#" & CStr(ErrorValue)
        End If

End Function
