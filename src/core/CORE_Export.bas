Attribute VB_Name = "CORE_Export"
'==============================================================================
' MODULE: CORE_Export
'------------------------------------------------------------------------------
' PURPOSE
'   Write the evidence report of a run as CSV text and UTF-8 bytes (#45),
'   the same on every computer: fields separated by ";", numbers with a
'   decimal point and no thousands separator, logical values as TRUE and
'   FALSE, error values by their Excel names, text quoted only when it
'   holds a separator, a quote or a line break, and lines ended by CR LF.
'   Each line starts with its section: RUN, PARAMS, SF_TABLE, FX_TABLE,
'   RESULTS, RECONCILIATION, TRADE_RECONCILIATION, CHECKS or LIMITATIONS.
'
' PUBLIC SURFACE
'   None outside this VBA project. NumberText, CsvField, CsvLine,
'   TableLines, LimitationLines and Utf8Bytes are Public for M_Main and
'   CORE_Reconcile; Option Private Module keeps them off the supported
'   external surface.
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
'   2026-10-16
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
    ByVal x As Double) _
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
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim s   As String    'The number as Str$ writes it

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        s = Trim$(Str$(x))
        If Left$(s, 1) = "." Then
            s = "0" & s
        ElseIf Left$(s, 2) = "-." Then
            s = "-0" & Mid$(s, 2)
        End If
        NumberText = s

End Function


Public Function CsvField( _
    ByVal v As Variant) _
    As String
'
'==============================================================================
'                                   CsvField
'------------------------------------------------------------------------------
' PURPOSE
'   Write one cell value (Value2) as a CSV field.
'
' RETURNS
'   "" for an empty cell; NumberText for a number; TRUE or FALSE; the Excel
'   name of an error value (#N/A, #VALUE!, ...); text as it is, in double
'   quotes with each quote doubled when it holds ";", a quote or a line
'   break.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim s   As String    'Text of the value

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        Select Case VarType(v)
            Case vbEmpty, vbNull
                CsvField = ""
            Case vbBoolean
                CsvField = IIf(v, "TRUE", "FALSE")
            Case vbError
                CsvField = ErrorName(v)
            Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte, vbDate
                CsvField = NumberText(CDbl(v))
            Case Else
                s = CStr(v)
                If InStr(s, SEP) > 0 Or InStr(s, QUOTE_CHR) > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
                    s = QUOTE_CHR & Replace(s, QUOTE_CHR, QUOTE_CHR & QUOTE_CHR) & QUOTE_CHR
                End If
                CsvField = s
        End Select

End Function


Public Function CsvLine( _
    ByVal section As String, _
    ByVal fields As Variant) _
    As String
'
'==============================================================================
'                                   CsvLine
'------------------------------------------------------------------------------
' PURPOSE
'   Write one line: the section, then each field, ended by CR LF.
'
' INPUTS
'   section: RUN, PARAMS, ...
'   fields: a 1-D array of values.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim parts()   As String    'Section and fields
    Dim i         As Long      'Field

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        ReDim parts(0 To UBound(fields) - LBound(fields) + 1)
        parts(0) = CsvField(section)
        For i = LBound(fields) To UBound(fields)
            parts(i - LBound(fields) + 1) = CsvField(fields(i))
        Next i
        CsvLine = Join(parts, SEP) & vbCrLf

End Function


Public Function TableLines( _
    ByVal section As String, _
    ByVal block As Variant) _
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
'   section: the section of every line.
'   block: a 1-based 2-D array, as Range.Value2 gives one; "" when it is
'   not an array.
'
' UPDATED
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim lines()   As String     'One line per row
    Dim fields()  As Variant    'Fields of one row
    Dim r         As Long       'Row
    Dim c         As Long       'Column
    Dim nCols     As Long       'Columns

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        If Not IsArray(block) Then
            Exit Function
        End If
        nCols = UBound(block, 2) - LBound(block, 2) + 1
        ReDim lines(LBound(block, 1) To UBound(block, 1))
        ReDim fields(1 To nCols)
        For r = LBound(block, 1) To UBound(block, 1)
            For c = 1 To nCols
                fields(c) = block(r, LBound(block, 2) + c - 1)
            Next c
            lines(r) = CsvLine(section, fields)
        Next r
        TableLines = Join(lines, "")

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
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim notes   As Variant    'One limitation per line
    Dim i       As Long       'Limitation
    Dim s       As String     'The lines

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        notes = Array( _
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
        For i = LBound(notes) To UBound(notes)
            s = s & CsvLine("LIMITATIONS", Array(i + 1, notes(i)))
        Next i
        LimitationLines = s

End Function


'
'------------------------------------------------------------------------------
'
'                                    BYTES
'
'------------------------------------------------------------------------------
'

Public Function Utf8Bytes( _
    ByVal text As String) _
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
'   2026-10-16
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim out()   As Byte     'Bytes written
    Dim n       As Long     'Bytes used
    Dim i       As Long     'Position in text
    Dim cp      As Long     'Code point
    Dim lo      As Long     'Low surrogate

'------------------------------------------------------------------------------
' ENCODE
'------------------------------------------------------------------------------
    'At most four bytes per code unit pair, three per unit otherwise.
        If Len(text) = 0 Then
            Utf8Bytes = out
            Exit Function
        End If
        ReDim out(0 To Len(text) * 3 - 1)
        i = 1
        Do While i <= Len(text)
            cp = AscW(Mid$(text, i, 1)) And &HFFFF&
            If cp >= &HD800& And cp <= &HDBFF& And i < Len(text) Then
                lo = AscW(Mid$(text, i + 1, 1)) And &HFFFF&
                If lo >= &HDC00& And lo <= &HDFFF& Then
                    cp = &H10000 + (cp - &HD800&) * &H400& + (lo - &HDC00&)
                    i = i + 1
                Else
                    cp = &HFFFD&
                End If
            ElseIf cp >= &HD800& And cp <= &HDFFF& Then
                cp = &HFFFD&
            End If
            If cp < &H80& Then
                out(n) = cp
                n = n + 1
            ElseIf cp < &H800& Then
                out(n) = &HC0& Or (cp \ &H40&)
                out(n + 1) = &H80& Or (cp And &H3F&)
                n = n + 2
            ElseIf cp < &H10000 Then
                out(n) = &HE0& Or (cp \ &H1000&)
                out(n + 1) = &H80& Or ((cp \ &H40&) And &H3F&)
                out(n + 2) = &H80& Or (cp And &H3F&)
                n = n + 3
            Else
                out(n) = &HF0& Or (cp \ &H40000)
                out(n + 1) = &H80& Or ((cp \ &H1000&) And &H3F&)
                out(n + 2) = &H80& Or ((cp \ &H40&) And &H3F&)
                out(n + 3) = &H80& Or (cp And &H3F&)
                n = n + 4
            End If
            i = i + 1
        Loop
        ReDim Preserve out(0 To n - 1)
        Utf8Bytes = out

End Function


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function ErrorName( _
    ByVal v As Variant) _
    As String
'
'==============================================================================
'                                  ErrorName
'------------------------------------------------------------------------------
' PURPOSE
'   Return the Excel name of an error value, the same in every language.
'
' UPDATED
'   2026-10-16
'==============================================================================
'
    'CLng of an error value raises, so the value is compared with each
    'error in turn; CStr gives "Error 2042" for one not listed.
        If v = CVErr(2000) Then
            ErrorName = "#NULL!"
        ElseIf v = CVErr(2007) Then
            ErrorName = "#DIV/0!"
        ElseIf v = CVErr(2015) Then
            ErrorName = "#VALUE!"
        ElseIf v = CVErr(2023) Then
            ErrorName = "#REF!"
        ElseIf v = CVErr(2029) Then
            ErrorName = "#NAME?"
        ElseIf v = CVErr(2036) Then
            ErrorName = "#NUM!"
        ElseIf v = CVErr(2042) Then
            ErrorName = "#N/A"
        Else
            ErrorName = "#" & CStr(v)
        End If

End Function
