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
'   section: RUN, PARAMS, SF_TABLE, FX_TABLE, HOLIDAY_TABLE, RESULTS, RECONCILIATION,
'   TRADE_RECONCILIATION, CHECKS or LIMITATIONS. Write the values of the
'   PDF of the formatted run report (#259) as text that does not depend on
'   the regional settings either, and group its Checks by code.
'
' PUBLIC SURFACE
'   None outside this VBA project. NumberText, CsvField, CsvLine,
'   TableLines, LimitationTexts, LimitationLines, Utf8Bytes and the
'   Report* functions, the text of the formatted run report (#259), are
'   Public for M_Main, CORE_Engine, CORE_Reconcile and the tests; Option
'   Private Module keeps them off the supported external surface.
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
'   2026-11-26
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


Public Function LimitationTexts() As Variant

'
'------------------------------------------------------------------------------
'                    LIST WHAT AN EXPORT OF A RUN DOES NOT SHOW
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the statements of the LIMITATIONS section of every evidence
'   report: what the report does and does not show
'
' WHY THIS EXISTS
'   The evidence CSV (#45) and the formatted run report (#259) state the
'   same limitations; one list keeps them the same
'
' INPUTS
'   None
'
' RETURNS
'   A zero-based array of the seven statements, in their order
'
' BEHAVIOR
'   Returns the fixed statements; nothing is read
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   ReportLimitationTexts adds the statement of the formatted report
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'The statements, one per element
        LimitationTexts = Array( _
            "The report describes one run of this workbook on the inputs fingerprinted in RUN; it is not a " & _
                "regulatory certification.", _
            "A comparison with a bank's figures is not independent validation of either.", _
            "Expected values classed illustrative show a behaviour and validate nothing.", _
            "Totals cover the VALID netting sets of one regime only; regimes are never added together.", _
            "Netting sets that are INVALID, INCOMPLETE or NO TRADES have no EAD and are listed on RESULTS " & _
                "and CHECKS.", _
            "Bank amounts are scaled to units as stated in RUN; ratios and times are not scaled.", _
            "Lines holding a run or reconciliation time change at every run; every other line depends only " & _
                "on the inputs.")

End Function


Public Function LimitationLines() As String

'
'------------------------------------------------------------------------------
'                     WRITE THE LIMITATIONS OF AN EVIDENCE CSV
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the fixed LIMITATIONS section of every evidence report: one
'   line per statement of LimitationTexts, numbered from 1
'
' WHY THIS EXISTS
'   The evidence CSV (#45) closes with what it does and does not show
'
' INPUTS
'   None
'
' RETURNS
'   The lines, each ended by CR LF
'
' BEHAVIOR
'   Writes "LIMITATIONS;n;statement" for each statement, in order
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   LimitationTexts, CsvLine
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Notes   As Variant    'One limitation per element
    Dim i       As Long       'Limitation
    Dim Lines   As String     'The lines

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'One numbered line per statement
        Notes = LimitationTexts()
        For i = LBound(Notes) To UBound(Notes)
            Lines = Lines & CsvLine("LIMITATIONS", Array(i + 1, Notes(i)))
        Next i
        LimitationLines = Lines

End Function


'
'------------------------------------------------------------------------------
'
'                                 REPORT TEXT
'
'------------------------------------------------------------------------------
'

Public Function ReportCellText( _
    ByVal Kind As String, _
    ByVal Value As Variant) _
    As String
'
'------------------------------------------------------------------------------
'                    WRITE A VALUE AS THE PDF OF A REPORT SHOWS IT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the text of a value of the formatted run report (#259) in a
'   column of the given kind, the same whatever the regional settings
'
' WHY THIS EXISTS
'   A PDF is read as it was written, so its numbers are text whose
'   separators do not follow the machine that wrote it.
'   tools/report.py (cell_text) is the reference, and
'   docs/workbook/report.json holds the examples both are checked against
'
' INPUTS
'   Kind
'     amount, ratio, count, date, datetime or text (report.json)
'   Value
'     A cell value (Value2): a number, a text, Empty, a logical value or an
'     error value
'
' RETURNS
'   "" for Empty and ""; an error value by its Excel name; TRUE or FALSE;
'   a text, also in a column of numbers, as it is; a number as its kind
'   writes it
'
' BEHAVIOR
'   - amount: two decimals, thousands grouped (ReportNumberText)
'   - ratio: six decimals, not grouped
'   - count: no decimal, not grouped
'   - date and datetime: ReportDateText and ReportDateTimeText
'   - text: the number with up to 15 significant digits (NumberText)
'
' ERROR POLICY
'   Raises ERR_REPORT_VALUE for an unknown kind, and the error of
'   ReportDateText for a date before 1 March 1900
'
' DEPENDENCIES
'   ReportNumberText, ReportDateText, ReportDateTimeText, NumberText,
'   ErrorName
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Export.ReportCellText"

    Dim Number          As Double    'Value as a number

'------------------------------------------------------------------------------
' WRITE A VALUE THAT IS NOT A NUMBER
'------------------------------------------------------------------------------
    'Blanks, logical values, error values and texts are written as they are
        Select Case VarType(Value)
            Case vbEmpty, vbNull
                ReportCellText = vbNullString
                Exit Function
            Case vbBoolean
                ReportCellText = IIf(Value, "TRUE", "FALSE")
                Exit Function
            Case vbError
                ReportCellText = ErrorName(Value)
                Exit Function
            Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte, vbDate
                Number = CDbl(Value)
            Case Else
                ReportCellText = CStr(Value)
                Exit Function
        End Select

'------------------------------------------------------------------------------
' WRITE A NUMBER BY ITS KIND
'------------------------------------------------------------------------------
    'Each kind has one invariant text
        Select Case Kind
            Case "amount"
                ReportCellText = ReportNumberText(Number, 2, True)
            Case "ratio"
                ReportCellText = ReportNumberText(Number, 6, False)
            Case "count"
                ReportCellText = ReportNumberText(Number, 0, False)
            Case "date"
                ReportCellText = ReportDateText(Number)
            Case "datetime"
                ReportCellText = ReportDateTimeText(Number)
            Case "text"
                ReportCellText = NumberText(Number)
            Case Else
                Err.Raise ERR_REPORT_VALUE, PROC_NAME, _
                    "Unknown kind of a report column: '" & Kind & "'"
        End Select

End Function


Public Function ReportNumberText( _
    ByVal Number As Double, _
    ByVal Decimals As Long, _
    ByVal Thousands As Boolean) _
    As String
'
'------------------------------------------------------------------------------
'                  WRITE A NUMBER WITH FIXED DECIMALS, ANY LOCALE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a number rounded to a number of decimals, with a point before
'   the decimals and, when asked, a comma between thousands
'
' WHY THIS EXISTS
'   Format$ follows the regional settings; the PDF of a report must not
'   (#259). The rounding works on the digits of the 15-significant-digit
'   text that NumberText writes, so that it gives what tools/report.py
'   gives: 2.675, a double just below 2.675, is written 2.68
'
' INPUTS
'   Number
'     Any finite number
'   Decimals
'     Decimals to keep, 0 or more
'   Thousands
'     True to put a comma between each group of three digits of the whole
'     part
'
' RETURNS
'   The text, such as 1,234,567.89; a value that rounds to zero has no
'   sign
'
' BEHAVIOR
'   - Writes |Number| with NumberText and reads its digits, its point and
'     its exponent
'   - Keeps the digits up to the last decimal and rounds half away from
'     zero on the next one, carrying to the left
'   - Groups the whole part and adds the sign of a value not rounded to
'     zero
'
' ERROR POLICY
'   None: every finite number has a text
'
' DEPENDENCIES
'   NumberText, IncrementDigits, GroupThousands
'
' NOTES
'   NumberText writes at most 15 significant digits, with an exponent
'   (1E-07, 1.5E+20) outside the range Str$ writes plainly
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Plain        As String    'The 15-digit text of |Number|
    Dim Mantissa     As String    'Its part before the exponent
    Dim Exponent     As Long      'Its power of ten; 0 when none
    Dim MarkAt       As Long      'Position of "E"; 0 when none
    Dim PointAt      As Long      'Position of the point; 0 when none
    Dim Digits       As String    'Every digit of the mantissa
    Dim WholeCount   As Long      'Digits before the point once the exponent applies
    Dim KeepCount    As Long      'Digits kept up to the last decimal
    Dim Kept         As String    'The kept digits, rounded
    Dim NextDigit    As Long      'The first digit dropped
    Dim WholePart    As String    'Digits before the point
    Dim Formatted    As String    'The text

'------------------------------------------------------------------------------
' READ THE DIGITS
'------------------------------------------------------------------------------
    'Split the 15-digit text into its digits and the place of its point
        Plain = NumberText(Abs(Number))
        MarkAt = VBA.InStr(1, Plain, "E", vbBinaryCompare)
        If MarkAt > 0 Then
            Exponent = CLng(VBA.Val(VBA.Mid$(Plain, MarkAt + 1)))
            Mantissa = VBA.Left$(Plain, MarkAt - 1)
        Else
            Mantissa = Plain
        End If
        PointAt = VBA.InStr(1, Mantissa, ".", vbBinaryCompare)
        If PointAt > 0 Then
            Digits = VBA.Left$(Mantissa, PointAt - 1) & VBA.Mid$(Mantissa, PointAt + 1)
            WholeCount = PointAt - 1 + Exponent
        Else
            Digits = Mantissa
            WholeCount = VBA.Len(Mantissa) + Exponent
        End If

'------------------------------------------------------------------------------
' ROUND HALF AWAY FROM ZERO
'------------------------------------------------------------------------------
    'Keep the digits up to the last decimal; the next digit decides
        KeepCount = WholeCount + Decimals
        If KeepCount < 0 Then
            Kept = "0"
            NextDigit = 0
        ElseIf KeepCount = 0 Then
            Kept = "0"
            NextDigit = CLng(VBA.Left$(Digits, 1))
        Else
            If VBA.Len(Digits) <= KeepCount Then
                Digits = Digits & VBA.String$(KeepCount - VBA.Len(Digits) + 1, "0")
            End If
            Kept = VBA.Left$(Digits, KeepCount)
            NextDigit = CLng(VBA.Mid$(Digits, KeepCount + 1, 1))
        End If
        If NextDigit >= 5 Then
            Kept = IncrementDigits(Kept)
        End If

'------------------------------------------------------------------------------
' WRITE THE TEXT
'------------------------------------------------------------------------------
    'Drop the leading zeros, keep one digit before the point
        Do While VBA.Len(Kept) > 1 And VBA.Left$(Kept, 1) = "0"
            Kept = VBA.Mid$(Kept, 2)
        Loop
        If VBA.Len(Kept) < Decimals + 1 Then
            Kept = VBA.String$(Decimals + 1 - VBA.Len(Kept), "0") & Kept
        End If
    'Group the whole part and add the decimals
        WholePart = VBA.Left$(Kept, VBA.Len(Kept) - Decimals)
        If Thousands Then
            WholePart = GroupThousands(WholePart)
        End If
        Formatted = WholePart
        If Decimals > 0 Then
            Formatted = Formatted & "." & VBA.Right$(Kept, Decimals)
        End If
    'A value rounded to zero has no sign
        If Number < 0# And VBA.Len(VBA.Replace(Kept, "0", vbNullString)) > 0 Then
            Formatted = "-" & Formatted
        End If
        ReportNumberText = Formatted

End Function


Private Function IncrementDigits( _
    ByVal Digits As String) _
    As String
'
'------------------------------------------------------------------------------
'                         ADD ONE TO A STRING OF DIGITS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a string of decimal digits plus one, carrying to the left:
'   "199" gives "200" and "99" gives "100"
'
' WHY THIS EXISTS
'   ReportNumberText rounds up on the digits of a text, not on a binary number
'
' INPUTS
'   Digits
'     A string of decimal digits, at least one
'
' RETURNS
'   The digits plus one, one digit longer when every digit was 9
'
' BEHAVIOR
'   Turns each 9 from the right into 0 and adds one to the first digit below 9
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result   As String    'The digits being changed
    Dim i        As Long      'Position, from the right
    Dim Digit    As Long      'Digit at the position

'------------------------------------------------------------------------------
' CARRY
'------------------------------------------------------------------------------
    'A 9 becomes 0 and carries; the first digit below 9 takes the one
        Result = Digits
        For i = VBA.Len(Result) To 1 Step -1
            Digit = CLng(VBA.Mid$(Result, i, 1))
            If Digit < 9 Then
                Mid$(Result, i, 1) = VBA.CStr(Digit + 1)
                IncrementDigits = Result
                Exit Function
            End If
            Mid$(Result, i, 1) = "0"
        Next i
        IncrementDigits = "1" & Result

End Function


Private Function GroupThousands( _
    ByVal WholePart As String) _
    As String
'
'------------------------------------------------------------------------------
'                        PUT A COMMA BETWEEN THOUSANDS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a string of digits with a comma before each group of three
'   from the right: "1234567" gives "1,234,567"
'
' WHY THIS EXISTS
'   The thousands of an amount in the PDF of a report (#259)
'
' INPUTS
'   WholePart
'     A string of decimal digits
'
' RETURNS
'   The digits with their commas
'
' BEHAVIOR
'   Takes three digits from the right at a time
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Rest      As String    'Digits not grouped yet
    Dim Grouped   As String    'Groups written so far

'------------------------------------------------------------------------------
' GROUP
'------------------------------------------------------------------------------
    'Take three digits from the right at a time
        Rest = WholePart
        Do While VBA.Len(Rest) > 3
            Grouped = "," & VBA.Right$(Rest, 3) & Grouped
            Rest = VBA.Left$(Rest, VBA.Len(Rest) - 3)
        Loop
        GroupThousands = Rest & Grouped

End Function


Public Function ReportDateText( _
    ByVal Serial As Double) _
    As String
'
'------------------------------------------------------------------------------
'                        WRITE A DATE AS YYYY-MM-DD
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the date of an Excel serial number as yyyy-mm-dd, any locale
'
' WHY THIS EXISTS
'   The dates of the PDF of a report (#259); tools/report.py (date_text)
'   is the reference
'
' INPUTS
'   Serial
'     A date of the 1900 date system, 61 (1 March 1900) or later; a time
'     of day is ignored
'
' RETURNS
'   The date, such as 2026-09-30
'
' BEHAVIOR
'   Reads year, month and day of the whole serial and writes them with
'   four, two and two digits
'
' ERROR POLICY
'   Raises ERR_REPORT_VALUE for a serial before 61, where VBA's and
'   Excel's date systems differ
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Export.ReportDateText"

    Dim TheDate         As Date    'The date

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'VBA and Excel agree from 1 March 1900 on
        If Serial < 61# Then
            Err.Raise ERR_REPORT_VALUE, PROC_NAME, _
                "A date before 1 March 1900 is not reported: " & NumberText(Serial)
        End If

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Year, month and day with fixed digits
        TheDate = CDate(Int(Serial))
        ReportDateText = VBA.Right$("000" & VBA.CStr(VBA.Year(TheDate)), 4) & "-" & _
                         VBA.Right$("0" & VBA.CStr(VBA.Month(TheDate)), 2) & "-" & _
                         VBA.Right$("0" & VBA.CStr(VBA.Day(TheDate)), 2)

End Function


Public Function ReportDateTimeText( _
    ByVal Serial As Double) _
    As String
'
'------------------------------------------------------------------------------
'                    WRITE A DATE AND TIME AS YYYY-MM-DD HH:MM:SS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the date and time of an Excel serial number to the second,
'   any locale
'
' WHY THIS EXISTS
'   The run time of a report (#259); tools/report.py (datetime_text) is
'   the reference
'
' INPUTS
'   Serial
'     A date and time of the 1900 date system, 61 or later
'
' RETURNS
'   The date and time, such as 2026-11-26 12:49:00
'
' BEHAVIOR
'   - Reads the 15-digit text of the serial (NumberText)
'   - Multiplies it by 86,400 in Decimal arithmetic, from its digits
'     only, so that no regional separator is read
'   - Rounds half up to the second; a time that rounds to midnight moves
'     to the next day
'
' ERROR POLICY
'   Raises ERR_REPORT_VALUE for a serial before 61
'
' DEPENDENCIES
'   NumberText, ReportDateText
'
' NOTES
'   CDec reads a string with the regional decimal separator, so it is
'   only given strings of digits
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME       As String = "CORE_Export.ReportDateTimeText"
    Const DAY_SECONDS     As Long = 86400

    Dim Plain             As String     'The 15-digit text of the serial
    Dim PointAt           As Long       'Position of its point; 0 when none
    Dim FractionDigits    As String     'Digits after the point
    Dim Seconds           As Variant    'Seconds from day 0, as a Decimal
    Dim WholeDays         As Variant    'Whole days of the rounded seconds
    Dim Rest              As Long       'Seconds within the day

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'VBA and Excel agree from 1 March 1900 on
        If Serial < 61# Then
            Err.Raise ERR_REPORT_VALUE, PROC_NAME, _
                "A date before 1 March 1900 is not reported: " & NumberText(Serial)
        End If

'------------------------------------------------------------------------------
' COUNT THE SECONDS
'------------------------------------------------------------------------------
    'Whole days and the fraction, from the digits of the 15-digit text
        Plain = NumberText(Serial)
        PointAt = VBA.InStr(1, Plain, ".", vbBinaryCompare)
        If PointAt > 0 Then
            FractionDigits = VBA.Mid$(Plain, PointAt + 1)
            Seconds = CDec(VBA.Left$(Plain, PointAt - 1)) * DAY_SECONDS + _
                      CDec(FractionDigits) * DAY_SECONDS / CDec("1" & VBA.String$(VBA.Len(FractionDigits), "0"))
        Else
            Seconds = CDec(Plain) * DAY_SECONDS
        End If
    'Round half up to the second
        Seconds = Int(Seconds + CDec(1) / CDec(2))

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'The day, then the time of day
        WholeDays = Int(Seconds / DAY_SECONDS)
        Rest = CLng(Seconds - WholeDays * DAY_SECONDS)
        ReportDateTimeText = ReportDateText(CDbl(WholeDays)) & " " & _
                             VBA.Right$("0" & VBA.CStr(Rest \ 3600), 2) & ":" & _
                             VBA.Right$("0" & VBA.CStr((Rest Mod 3600) \ 60), 2) & ":" & _
                             VBA.Right$("0" & VBA.CStr(Rest Mod 60), 2)

End Function


Public Function ReportFormatText( _
    ByVal Value As Variant, _
    ByVal NumberFormat As String) _
    As String
'
'------------------------------------------------------------------------------
'                WRITE A VALUE AS ITS NUMBER FORMAT SHOWS IT, ANY LOCALE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the text of a value as the number format of its cell would
'   show it, with a point before the decimals and a comma between
'   thousands whatever the regional settings
'
' WHY THIS EXISTS
'   The Dashboard's figures in the PDF of a report (#259): each block
'   keeps the format its cell has on the sheet. tools/report.py
'   (format_text) is the reference
'
' INPUTS
'   Value
'     A cell value (Value2)
'   NumberFormat
'     The cell's number format, in Excel's invariant codes
'
' RETURNS
'   The text; a value that is not a number as ReportCellText writes it
'
' BEHAVIOR
'   - A format with a year or a day code is a date, with an hour code a
'     date and time
'   - General, text and an empty format: the 15-digit text (NumberText)
'   - Otherwise the decimals are the digit placeholders after the point
'     of the first section, a comma before the point groups thousands and
'     a percent sign multiplies by 100 and is written after the number
'
' ERROR POLICY
'   Raises the error of ReportDateText for a date before 1 March 1900
'
' DEPENDENCIES
'   ReportCellText, ReportNumberText, ReportDateText, ReportDateTimeText,
'   NumberText
'
' NOTES
'   Literal text, colours and conditions of a format are not reproduced
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Section      As String     'First section of the format, in capitals
    Dim PointAt      As Long       'Position of its point; 0 when none
    Dim Decimals     As Long       'Digit placeholders after the point
    Dim i            As Long       'Position after the point
    Dim Percent      As Boolean    'The format ends in a percent sign
    Dim Thousands    As Boolean    'The format groups thousands
    Dim Number       As Double     'Value as a number

'------------------------------------------------------------------------------
' WRITE A VALUE THAT IS NOT A NUMBER
'------------------------------------------------------------------------------
    'Only numbers depend on the format
        Select Case VarType(Value)
            Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte, vbDate
                Number = CDbl(Value)
            Case Else
                ReportFormatText = ReportCellText("text", Value)
                Exit Function
        End Select

'------------------------------------------------------------------------------
' READ THE FORMAT
'------------------------------------------------------------------------------
    'Its first section decides a positive and a negative number alike
        Section = VBA.UCase$(NumberFormat)
        If VBA.InStr(1, Section, ";", vbBinaryCompare) > 0 Then
            Section = VBA.Left$(Section, VBA.InStr(1, Section, ";", vbBinaryCompare) - 1)
        End If
        If VBA.InStr(1, Section, "YY", vbBinaryCompare) > 0 Or VBA.InStr(1, Section, "DD", vbBinaryCompare) > 0 Then
            If VBA.InStr(1, Section, "H", vbBinaryCompare) > 0 Then
                ReportFormatText = ReportDateTimeText(Number)
            Else
                ReportFormatText = ReportDateText(Number)
            End If
            Exit Function
        End If
        If VBA.Len(Section) = 0 Or Section = "GENERAL" Or Section = "@" Then
            ReportFormatText = NumberText(Number)
            Exit Function
        End If
    'Decimals, grouping and percent
        PointAt = VBA.InStr(1, Section, ".", vbBinaryCompare)
        If PointAt > 0 Then
            For i = PointAt + 1 To VBA.Len(Section)
                If VBA.Mid$(Section, i, 1) = "0" Or VBA.Mid$(Section, i, 1) = "#" Then
                    Decimals = Decimals + 1
                Else
                    Exit For
                End If
            Next i
            Thousands = VBA.InStr(1, VBA.Left$(Section, PointAt), ",", vbBinaryCompare) > 0
        Else
            Thousands = VBA.InStr(1, Section, ",", vbBinaryCompare) > 0
        End If
        Percent = VBA.InStr(1, Section, "%", vbBinaryCompare) > 0

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'A percentage is the number times 100, then the sign
        If Percent Then
            ReportFormatText = ReportNumberText(Number * 100#, Decimals, Thousands) & "%"
        Else
            ReportFormatText = ReportNumberText(Number, Decimals, Thousands)
        End If

End Function


Public Function ReportMeasureKind( _
    ByVal Measure As String, _
    ByVal RatioMeasures As String) _
    As String
'
'------------------------------------------------------------------------------
'                    CHOOSE THE FORMAT OF A RECONCILED FIGURE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the kind of the figures of a reconciliation row: ratio for a
'   measure that is not an amount, amount for every other
'
' WHY THIS EXISTS
'   A multiplier or a maturity written with two decimals would hide a
'   difference (#259); tools/report.py (measure_kind) is the reference
'
' INPUTS
'   Measure
'     The Measure of the row, as the reconciliation sheet holds it
'   RatioMeasures
'     The measures that are not amounts, separated by "|"
'
' RETURNS
'   "ratio" or "amount"
'
' BEHAVIOR
'   Compares the measure with each name of the list, case included
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   M_Main passes the lists of docs/workbook/report.json (measure_ratios)
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'A name of the list is a ratio
        If VBA.InStr(1, "|" & RatioMeasures & "|", "|" & Measure & "|", vbBinaryCompare) > 0 Then
            ReportMeasureKind = "ratio"
        Else
            ReportMeasureKind = "amount"
        End If

End Function


Public Function ReportCodeOrder( _
    ByVal CheckRows As Variant, _
    ByVal SeverityColumn As Long, _
    ByVal CodeColumn As Long, _
    ByRef Summary As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'                         GROUP THE ROWS OF CHECKS BY CODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the order in which the report lists the rows of Checks, and
'   the summary of their codes
'
' WHY THIS EXISTS
'   The report gives Checks by code (#259, owner's decision of
'   2026-11-26): a row per code, then the rows grouped by code.
'   tools/report.py (checks_by_code, checks_in_code_order) is the reference
'
' INPUTS
'   CheckRows
'     The rows of Checks, a 2-D array with bounds from 1 as Range.Value2
'     gives it; Empty when there is none
'   SeverityColumn
'     Column of the severity in CheckRows
'   CodeColumn
'     Column of the code in CheckRows
'   Summary
'     Written: a 2-D array, one row per code, of the code, its highest
'     severity and its number of rows; Empty when there is no row
'
' RETURNS
'   A 1-D array of the row numbers of CheckRows, from 1, in the order of the
'   summary and, within a code, as on the sheet; Empty when there is no
'   row
'
' BEHAVIOR
'   - A blank code, also one of spaces, is the code "(none)"
'   - The highest severity of a code is the first of ERROR, WARNING and
'     INFO among its rows; another severity is the lowest and is written
'     blank
'   - Codes are ordered by that severity, then by code in ordinal order
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   ReportCheckCode, ErrorName
'
' NOTES
'   Checks holds few codes, so the codes are sorted by insertion
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RowCount      As Long      'Rows of Checks
    Dim RowBase       As Long      'Lower bound of the rows, less one
    Dim ColumnBase    As Long      'Lower bound of the columns, less one
    Dim Codes()       As String    'Distinct codes, then sorted
    Dim Ranks()       As Long      'Highest severity of each code: 0 ERROR to 3 other
    Dim Counts()      As Long      'Rows of each code
    Dim CodeCount     As Long      'Distinct codes
    Dim Code          As String    'Code of a row
    Dim Rank          As Long      'Severity of a row
    Dim Severity      As String    'Severity text of a row
    Dim Order()       As Long      'Row numbers in the order of the report
    Dim Filled        As Long      'Row numbers written to Order
    Dim r             As Long      'Row of Checks
    Dim k             As Long      'Code
    Dim j             As Long      'Code compared while sorting
    Dim HeldCode      As String    'Code moved while sorting
    Dim HeldRank      As Long      'Its severity
    Dim HeldCount     As Long      'Its rows
    Dim Result        As Variant   'The summary

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'No row, no code
        Summary = Empty
        If Not IsArray(CheckRows) Then
            ReportCodeOrder = Empty
            Exit Function
        End If
        RowCount = UBound(CheckRows, 1) - LBound(CheckRows, 1) + 1
        RowBase = LBound(CheckRows, 1) - 1
        ColumnBase = LBound(CheckRows, 2) - 1

'------------------------------------------------------------------------------
' COUNT THE CODES
'------------------------------------------------------------------------------
    'Each code once, with its highest severity and its rows
        ReDim Codes(1 To RowCount)
        ReDim Ranks(1 To RowCount)
        ReDim Counts(1 To RowCount)
        For r = 1 To RowCount
            Code = ReportCheckCode(CheckRows(RowBase + r, ColumnBase + CodeColumn))
            Severity = ReportCellText("text", CheckRows(RowBase + r, ColumnBase + SeverityColumn))
            Severity = VBA.UCase$(VBA.Trim$(Severity))
            Select Case Severity
                Case "ERROR"
                    Rank = 0
                Case "WARNING"
                    Rank = 1
                Case "INFO"
                    Rank = 2
                Case Else
                    Rank = 3
            End Select
            For k = 1 To CodeCount
                If Codes(k) = Code Then Exit For
            Next k
            If k > CodeCount Then
                CodeCount = CodeCount + 1
                Codes(CodeCount) = Code
                Ranks(CodeCount) = Rank
            ElseIf Rank < Ranks(k) Then
                Ranks(k) = Rank
            End If
            Counts(k) = Counts(k) + 1
        Next r

'------------------------------------------------------------------------------
' SORT THE CODES
'------------------------------------------------------------------------------
    'By severity, then by code in ordinal order
        For k = 2 To CodeCount
            HeldCode = Codes(k)
            HeldRank = Ranks(k)
            HeldCount = Counts(k)
            j = k - 1
            Do While j >= 1
                If Ranks(j) < HeldRank Then Exit Do
                If Ranks(j) = HeldRank And VBA.StrComp(Codes(j), HeldCode, vbBinaryCompare) <= 0 Then Exit Do
                Codes(j + 1) = Codes(j)
                Ranks(j + 1) = Ranks(j)
                Counts(j + 1) = Counts(j)
                j = j - 1
            Loop
            Codes(j + 1) = HeldCode
            Ranks(j + 1) = HeldRank
            Counts(j + 1) = HeldCount
        Next k

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'The summary, one row per code
        ReDim Result(1 To CodeCount, 1 To 3)
        For k = 1 To CodeCount
            Result(k, 1) = Codes(k)
            Result(k, 2) = VBA.Choose(Ranks(k) + 1, "ERROR", "WARNING", "INFO", vbNullString)
            Result(k, 3) = Counts(k)
        Next k
        Summary = Result
    'The rows of each code in the order of the summary
        ReDim Order(1 To RowCount)
        For k = 1 To CodeCount
            For r = 1 To RowCount
                If ReportCheckCode(CheckRows(RowBase + r, ColumnBase + CodeColumn)) = Codes(k) Then
                    Filled = Filled + 1
                    Order(Filled) = r
                End If
            Next r
        Next k
        ReportCodeOrder = Order

End Function


Private Function ReportCheckCode( _
    ByVal Value As Variant) _
    As String
'
'------------------------------------------------------------------------------
'                            READ THE CODE OF A CHECK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the code of a row of Checks, trimmed; "(none)" for a blank one
'
' WHY THIS EXISTS
'   ReportCodeOrder compares the codes of the rows twice and must read them alike
'
' INPUTS
'   Value
'     The Code cell of a row of Checks (Value2)
'
' RETURNS
'   The code, or "(none)"
'
' BEHAVIOR
'   Writes the value as text (ReportCellText) and trims it
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   ReportCellText
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Code   As String    'The code as written

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'A blank code belongs to "(none)"
        Code = VBA.Trim$(ReportCellText("text", Value))
        If VBA.Len(Code) = 0 Then
            Code = "(none)"
        End If
        ReportCheckCode = Code

End Function


Public Function ReportOmittedNote( _
    ByVal Omitted As Long) _
    As String
'
'------------------------------------------------------------------------------
'                    SAY HOW MANY ROWS THE PDF LEAVES OUT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the line under a reconciliation in the PDF of a report that
'   says how many MATCH rows it leaves out; "" when none
'
' WHY THIS EXISTS
'   The PDF shows the rows that are not MATCH (#259, owner's decision of
'   2026-11-26), and an omission is never silent. tools/report.py
'   (omitted_note) is the reference
'
' INPUTS
'   Omitted
'     Rows left out, 0 or more
'
' RETURNS
'   The line, or "" for 0
'
' BEHAVIOR
'   Writes the number, in the singular for one
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Nothing to say when every row is shown
        If Omitted = 0 Then
            ReportOmittedNote = vbNullString
        ElseIf Omitted = 1 Then
            ReportOmittedNote = "1 row with the result MATCH is left out; the workbook version of the report " & _
                                "has every row."
        Else
            ReportOmittedNote = VBA.CStr(Omitted) & " rows with the result MATCH are left out; the workbook " & _
                                "version of the report has every row."
        End If

End Function


Public Function ReportLimitationTexts() As Variant

'
'------------------------------------------------------------------------------
'                      LIST THE LIMITATIONS OF A RUN REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the statements of the Limitations sheet of a formatted run
'   report: those of the evidence CSV and one more
'
' WHY THIS EXISTS
'   The report is a readable view of the run, and the CSV of Export
'   evidence its machine-readable record (#259); docs/workbook/report.json
'   counts eight statements
'
' INPUTS
'   None
'
' RETURNS
'   A zero-based array of the eight statements
'
' BEHAVIOR
'   Copies LimitationTexts and adds the report's statement
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   LimitationTexts
'
' NOTES
'   None
'
' UPDATED
'   2026-11-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Notes   As Variant    'The statements of the evidence CSV
    Dim Texts   As Variant    'The statements of the report
    Dim i       As Long       'Statement

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'The evidence CSV's statements, then the report's
        Notes = LimitationTexts()
        ReDim Texts(0 To UBound(Notes) - LBound(Notes) + 1)
        For i = LBound(Notes) To UBound(Notes)
            Texts(i - LBound(Notes)) = Notes(i)
        Next i
        Texts(UBound(Texts)) = "This report is a readable view of the run; the CSV written by Export evidence is " & _
                               "its machine-readable record."
        ReportLimitationTexts = Texts

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
