Attribute VB_Name = "CORE_Util"
'==============================================================================
' MODULE: CORE_Util
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the small helpers the engine uses everywhere: tolerant conversion
'   of cell values to text, numbers, dates and Booleans; keyed lookups built
'   on a VBA Collection; asset-class code translation; and the sheet helpers
'   that read parameters and input columns and write output blocks.
'
' PUBLIC SURFACE
'   None outside this VBA project. Every procedure is Public for in-project use
'   by CORE_Engine and M_Main; Option Private Module keeps them off the supported
'   external surface.
'
' DEPENDENCIES
'   CORE_Config for sheet names, layout constants and asset-class numbers. Keyed
'   lookups use a VBA Collection, so no Scripting.Dictionary reference is
'   needed and the code also runs on Excel for Mac.
'
' STATE OWNERSHIP
'   Stateless. The sheet helpers read and write the worksheets of ThisWorkbook
'   that their caller names; they change no Excel application setting.
'
' ERROR POLICY
'   Conversions never raise: an error value, blank or unusable input returns
'   the documented default. KeyIndex and NamedCell contain the
'   expected lookup errors and report "not found" instead. Every other
'   error, for example a missing worksheet, propagates to the caller.
'
' KNOWN DEVIATION
'   This module lives in src/core but reads and writes worksheets, which the
'   repository structure reserves for src/workbook. See the known deviations
'   in docs/REPOSITORY_STRUCTURE.md.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' UPDATED
'   2026-10-12
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the helpers inside this project.
    Option Explicit
    Option Private Module


'
'------------------------------------------------------------------------------
'
'                                 CONVERSIONS
'
'------------------------------------------------------------------------------
'

Public Function SafeStr( _
    ByVal v As Variant) _
    As String
'
'==============================================================================
'                                   SafeStr
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to trimmed text without failing on error values.
'
' INPUTS
'   v: any cell value, including #N/A, #VALUE! and other error values.
'
' RETURNS
'   The trimmed text of v; an empty string for an error value, Empty or Null.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'CStr raises on an error value, so error values, Empty and Null are
    'mapped to an empty string before converting.
        If IsError(v) Then
            SafeStr = ""
        ElseIf IsEmpty(v) Or IsNull(v) Then
            SafeStr = ""
        Else
            SafeStr = Trim$(CStr(v))
        End If

End Function


Public Function UTxt( _
    ByVal v As Variant) _
    As String
'
'==============================================================================
'                                     UTxt
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to trimmed upper-case text, so that codes such as
'   "long", "Long" and " LONG " compare equal.
'
' INPUTS
'   v: any cell value.
'
' RETURNS
'   UCase$ of SafeStr(v).
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
        UTxt = UCase$(SafeStr(v))

End Function


Public Function IsBlankCell( _
    ByVal v As Variant) _
    As Boolean
'
'==============================================================================
'                                 IsBlankCell
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a cell value is empty or contains only spaces.
'
' INPUTS
'   v: any cell value. An error value counts as blank.
'
' RETURNS
'   True when SafeStr(v) is empty.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
        IsBlankCell = (Len(SafeStr(v)) = 0)

End Function


Public Function IsNum( _
    ByVal v As Variant) _
    As Boolean
'
'==============================================================================
'                                    IsNum
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a cell value can be used as a number.
'
' INPUTS
'   v: any cell value.
'
' RETURNS
'   False for an error value, Empty or a blank string; otherwise IsNumeric(v).
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
    'IsNumeric alone treats Empty as numeric, so Empty and blank strings are
    'rejected first.
        If IsError(v) Or IsEmpty(v) Then
            IsNum = False
        ElseIf VarType(v) = vbString Then
            IsNum = (Len(Trim$(v)) > 0) And IsNumeric(v)
        Else
            IsNum = IsNumeric(v)
        End If

End Function


Public Function ToDbl( _
    ByVal v As Variant, _
    Optional ByVal dflt As Double = 0#) _
    As Double
'
'==============================================================================
'                                    ToDbl
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to a Double, falling back to a default.
'
' INPUTS
'   v: any cell value.
'   dflt: value returned when v is not a usable number; 0 by default.
'
' RETURNS
'   CDbl(v) when IsNum(v); otherwise dflt.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
        If IsNum(v) Then
            ToDbl = CDbl(v)
        Else
            ToDbl = dflt
        End If

End Function


Public Function ToSerial( _
    ByVal v As Variant) _
    As Double
'
'==============================================================================
'                                   ToSerial
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to a date serial number.
'
' INPUTS
'   v: a date, a number already holding a serial, or text in the ISO form
'      yyyy-mm-dd. Other text, such as 31/03/2027, 03/31/2027 or a serial
'      written as text, is not read: its meaning depends on the Windows
'      regional settings (#40).
'
' RETURNS
'   The date serial as a Double; -1 when v is missing or not a date.
'
' UPDATED
'   2026-10-12
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim s     As String    'Trimmed text of v
    Dim y     As Long      'Year of an ISO text date
    Dim m     As Long      'Month of an ISO text date
    Dim d     As Long      'Day of an ISO text date
    Dim dt    As Date      'The ISO text date, built from y, m and d

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'Start from the "missing" marker and replace it only when v is usable.
    'Text is tested before numbers: IsNum accepts numeric text such as
    '"47756", which is not a yyyy-mm-dd date.
        ToSerial = -1#
        If IsError(v) Or IsEmpty(v) Then
            Exit Function
        End If
        If VarType(v) = vbDate Then
            ToSerial = CDbl(v)
        ElseIf VarType(v) = vbString Then
            s = Trim$(v)
            If Not s Like "####-##-##" Then
                Exit Function
            End If

    'Only a calendar date is read: 2027-02-30 is not rolled into March.
            y = CLng(Left$(s, 4))
            m = CLng(Mid$(s, 6, 2))
            d = CLng(Right$(s, 2))
            If y < 1900 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then
                Exit Function
            End If
            dt = DateSerial(y, m, d)
            If Month(dt) = m And Day(dt) = d Then
                ToSerial = CDbl(dt)
            End If
        ElseIf IsNum(v) Then
            ToSerial = CDbl(v)
        End If

End Function


Public Function ToBool( _
    ByVal v As Variant, _
    Optional ByVal dflt As Boolean = False) _
    As Boolean
'
'==============================================================================
'                                    ToBool
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a Y/N style cell value to a Boolean.
'
' INPUTS
'   v: a Boolean, or text such as Y, YES, TRUE, 1, N, NO, FALSE, 0. The
'      Italian VERO, SI and FALSO are also accepted.
'   dflt: value returned for an error value, Empty or unrecognised text.
'
' RETURNS
'   True or False as read from v; otherwise dflt.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim s   As String    'Upper-case trimmed text of v

'------------------------------------------------------------------------------
' HANDLE NON-TEXT VALUES
'------------------------------------------------------------------------------
    'Error values and Empty take the default; a real Boolean is returned
    'as it is.
        If IsError(v) Or IsEmpty(v) Then
            ToBool = dflt
            Exit Function
        End If
        If VarType(v) = vbBoolean Then
            ToBool = v
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ TEXT
'------------------------------------------------------------------------------
        s = UTxt(v)
        Select Case s
            Case "Y", "YES", "TRUE", "1", "VERO", "SI"
                ToBool = True
            Case "N", "NO", "FALSE", "0", "FALSO"
                ToBool = False
            Case Else
                ToBool = dflt
        End Select

End Function


Public Function TryBool( _
    ByVal v As Variant, _
    ByVal dflt As Boolean, _
    ByRef result As Boolean) _
    As Boolean
'
'==============================================================================
'                                   TryBool
'------------------------------------------------------------------------------
' PURPOSE
'   Read a Y/N style cell value strictly: unlike ToBool, an unrecognised
'   value is reported instead of silently taking the default (#35).
'
' INPUTS
'   v: a Boolean, blank, or text such as Y, YES, TRUE, 1, N, NO, FALSE, 0
'      (and the Italian VERO, SI, FALSO).
'   dflt: value used when v is blank.
'
' RETURNS
'   True with result set when v is blank or recognised; False for an error
'   value or unrecognised text, with result left as dflt.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        result = dflt
        If IsError(v) Then
            Exit Function
        End If
        If VarType(v) = vbBoolean Then
            result = v
            TryBool = True
            Exit Function
        End If
        Select Case UTxt(v)
            Case ""
                TryBool = True
            Case "Y", "YES", "TRUE", "1", "VERO", "SI"
                result = True
                TryBool = True
            Case "N", "NO", "FALSE", "0", "FALSO"
                result = False
                TryBool = True
        End Select

End Function


Public Function TryDbl( _
    ByVal v As Variant, _
    ByVal dflt As Double, _
    ByRef result As Double) _
    As Boolean
'
'==============================================================================
'                                    TryDbl
'------------------------------------------------------------------------------
' PURPOSE
'   Read a numeric cell value strictly: unlike ToDbl, text that is not a
'   number is reported instead of silently taking the default (#35).
'
' INPUTS
'   v: a number, blank, or anything else.
'   dflt: value used when v is blank.
'
' RETURNS
'   True with result set when v is blank or a usable number; False for an
'   error value or non-numeric text, with result left as dflt.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        result = dflt
        If IsError(v) Then
            Exit Function
        End If
        If IsBlankCell(v) Then
            TryDbl = True
        ElseIf IsNum(v) Then
            result = CDbl(v)
            TryDbl = True
        End If

End Function


Public Function Max2( _
    ByVal a As Double, _
    ByVal b As Double) _
    As Double
'
'==============================================================================
'                                     Max2
'------------------------------------------------------------------------------
' PURPOSE
'   Return the larger of two numbers.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If a > b Then
            Max2 = a
        Else
            Max2 = b
        End If

End Function


Public Function Min2( _
    ByVal a As Double, _
    ByVal b As Double) _
    As Double
'
'==============================================================================
'                                     Min2
'------------------------------------------------------------------------------
' PURPOSE
'   Return the smaller of two numbers.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If a < b Then
            Min2 = a
        Else
            Min2 = b
        End If

End Function


'
'------------------------------------------------------------------------------
'
'                                KEYED LOOKUPS
'
'------------------------------------------------------------------------------
'

Public Function KeyIndex( _
    ByVal col As Collection, _
    ByVal key As String) _
    As Long
'
'==============================================================================
'                                   KeyIndex
'------------------------------------------------------------------------------
' PURPOSE
'   Look up the array index stored under a key.
'
' INPUTS
'   col: a Collection filled by KeyAdd.
'   key: lookup key; compared case-insensitively.
'
' RETURNS
'   The stored 1-based index; 0 when the key is not present.
'
' ERROR POLICY
'   Collection.Item raises for a missing key. That single expected error is
'   cleared and reported as 0.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim v   As Variant    'Value stored under the key

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
    'Keys are stored upper-case by KeyAdd, so the lookup upper-cases too.
        On Error Resume Next
        v = col.Item(UCase$(key))
        If Err.Number <> 0 Then
            Err.Clear
            KeyIndex = 0
        Else
            KeyIndex = CLng(v)
        End If
        On Error GoTo 0

End Function


Public Sub KeyAdd( _
    ByVal col As Collection, _
    ByVal key As String, _
    ByVal idx As Long)
'
'==============================================================================
'                                    KeyAdd
'------------------------------------------------------------------------------
' PURPOSE
'   Store an array index under a key.
'
' INPUTS
'   col: the Collection used as the index.
'   key: lookup key; stored upper-case.
'   idx: 1-based position of the item in its array.
'
' ERROR POLICY
'   Raises the Collection's duplicate-key error if the key already exists;
'   callers check with KeyIndex first.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' STORE
'------------------------------------------------------------------------------
        col.Add idx, UCase$(key)

End Sub


'
'------------------------------------------------------------------------------
'
'                                ASSET CLASSES
'
'------------------------------------------------------------------------------
'

Public Function ACIndex( _
    ByVal code As String) _
    As Long
'
'==============================================================================
'                                   ACIndex
'------------------------------------------------------------------------------
' PURPOSE
'   Translate an asset-class code into its internal number.
'
' INPUTS
'   code: IR, FX, CR, EQ, CO or OT, in any case.
'
' RETURNS
'   The matching AC_ constant from CORE_Config; 0 for an unknown code.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' TRANSLATE
'------------------------------------------------------------------------------
        Select Case UCase$(code)
            Case "IR"
                ACIndex = AC_IR
            Case "FX"
                ACIndex = AC_FX
            Case "CR"
                ACIndex = AC_CR
            Case "EQ"
                ACIndex = AC_EQ
            Case "CO"
                ACIndex = AC_CO
            Case "OT"
                ACIndex = AC_OT
            Case Else
                ACIndex = 0
        End Select

End Function


Public Function ACCode( _
    ByVal idx As Long) _
    As String
'
'==============================================================================
'                                    ACCode
'------------------------------------------------------------------------------
' PURPOSE
'   Translate an internal asset-class number back into its code.
'
' INPUTS
'   idx: an AC_ constant from CORE_Config.
'
' RETURNS
'   IR, FX, CR, EQ, CO or OT; "?" for any other number.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' TRANSLATE
'------------------------------------------------------------------------------
        Select Case idx
            Case AC_IR
                ACCode = "IR"
            Case AC_FX
                ACCode = "FX"
            Case AC_CR
                ACCode = "CR"
            Case AC_EQ
                ACCode = "EQ"
            Case AC_CO
                ACCode = "CO"
            Case AC_OT
                ACCode = "OT"
            Case Else
                ACCode = "?"
        End Select

End Function


'
'------------------------------------------------------------------------------
'
'                                SHEET HELPERS
'
'------------------------------------------------------------------------------
'

Public Function GetSheet( _
    ByVal sheetName As String) _
    As Worksheet
'
'==============================================================================
'                                   GetSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Return a worksheet of this workbook by tab name.
'
' INPUTS
'   sheetName: one of the SH_ constants in CORE_Config.
'
' ERROR POLICY
'   A missing sheet raises "Subscript out of range" to the caller.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN SHEET
'------------------------------------------------------------------------------
        Set GetSheet = ThisWorkbook.Worksheets(sheetName)

End Function


Public Function TableCell( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal col As Long) _
    As Range
'
'==============================================================================
'                                  TableCell
'------------------------------------------------------------------------------
' PURPOSE
'   Return the cell in a sheet row at a table column: column numbers count
'   from the first table column, right of the COL_OFFSET margin columns.
'
' INPUTS
'   ws: the worksheet.
'   rowNum: the sheet row.
'   col: the table column, 1 for the first.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        Set TableCell = ws.Cells(rowNum, col + COL_OFFSET)

End Function


Public Function NamedCell( _
    ByVal nameText As String) _
    As Range
'
'==============================================================================
'                                  NamedCell
'------------------------------------------------------------------------------
' PURPOSE
'   Return the top-left cell of a workbook name, or Nothing when the name
'   is missing or no longer refers to a cell.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim target   As Range    'The name's range

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Names(nameText) raises when no such name exists; target then stays
    'Nothing.
        On Error Resume Next
        Set target = ThisWorkbook.Names(nameText).RefersToRange
        Err.Clear
        On Error GoTo 0
        If Not target Is Nothing Then
            Set NamedCell = target.Cells(1, 1)
        End If

End Function


Public Function NamedRow( _
    ByVal nameText As String) _
    As Long
'
'==============================================================================
'                                   NamedRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the sheet row of a workbook name's top-left cell; 0 when the
'   name is missing or broken.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Range    'The name's top-left cell

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
        Set c = NamedCell(nameText)
        If Not c Is Nothing Then
            NamedRow = c.Row
        End If

End Function


Public Function TableHeaderRow( _
    ByVal sheetName As String) _
    As Long
'
'==============================================================================
'                                TableHeaderRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the header row of a sheet's table. NettingSets and Trades are
'   located through the workbook name of their first header cell
'   (NM_NS_HEADER, NM_TR_HEADER), which must be in column C of that sheet;
'   every other sheet has its headers on HEADER_ROW.
'
' RETURNS
'   The row; 0 when an input sheet's name is missing, broken, on another
'   sheet or not in column C.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim headerName   As String    'Workbook name of the header cell
    Dim c            As Range     'The named cell

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
        Select Case sheetName
            Case SH_NS
                headerName = NM_NS_HEADER
            Case SH_TRADES
                headerName = NM_TR_HEADER
            Case Else
                TableHeaderRow = HEADER_ROW
                Exit Function
        End Select
        Set c = NamedCell(headerName)
        If Not c Is Nothing Then
            If StrComp(c.Worksheet.Name, sheetName, vbTextCompare) = 0 And c.Column = COL_OFFSET + 1 Then
                TableHeaderRow = c.Row
            End If
        End If

End Function


Public Function TableFirstRow( _
    ByVal sheetName As String) _
    As Long
'
'==============================================================================
'                                TableFirstRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the first data row of a sheet's table: the row below its
'   headers (TableHeaderRow).
'
' RETURNS
'   The row; 0 when the header row cannot be located.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Header row

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
        r = TableHeaderRow(sheetName)
        If r > 0 Then
            TableFirstRow = r + 1
        End If

End Function


Public Function NamedTableCol( _
    ByVal nameText As String) _
    As Long
'
'==============================================================================
'                                NamedTableCol
'------------------------------------------------------------------------------
' PURPOSE
'   Return the column of a workbook name's top-left cell, counted as
'   TableCell counts table columns, so that the column can be passed to
'   TableCell and TableRange. Check NamedRow first: a missing name gives 0
'   here too.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Range    'The name's top-left cell

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
        Set c = NamedCell(nameText)
        If Not c Is Nothing Then
            NamedTableCol = c.Column - COL_OFFSET
        End If

End Function


Public Function ParamsHeaderRow() As Long
'
'==============================================================================
'                               ParamsHeaderRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the header row of the parameter rows on Params, from the workbook
'   name NM_PRM_CODE; 0 when the name is missing or broken.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        ParamsHeaderRow = NamedRow(NM_PRM_CODE)

End Function


Public Function ParamsCodeCol() As Long
'
'==============================================================================
'                                ParamsCodeCol
'------------------------------------------------------------------------------
' PURPOSE
'   Return the column of the parameter codes on Params, as a table column
'   for TableCell, from the workbook name NM_PRM_CODE.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        ParamsCodeCol = NamedTableCol(NM_PRM_CODE)

End Function


Public Function ParamsValueCol() As Long
'
'==============================================================================
'                                ParamsValueCol
'------------------------------------------------------------------------------
' PURPOSE
'   Return the column of the parameter values on Params, as a table column
'   for TableCell, from the workbook name NM_PRM_VALUE.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        ParamsValueCol = NamedTableCol(NM_PRM_VALUE)

End Function


Public Function TableRange( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal firstCol As Long, _
    ByVal lastRow As Long, _
    ByVal lastCol As Long) _
    As Range
'
'==============================================================================
'                                  TableRange
'------------------------------------------------------------------------------
' PURPOSE
'   Return the block between two sheet rows and two table columns, as
'   TableCell counts them.
'
' UPDATED
'   2026-10-11
'==============================================================================
'
        Set TableRange = ws.Range(TableCell(ws, firstRow, firstCol), TableCell(ws, lastRow, lastCol))

End Function


Public Function UsedLastRow( _
    ByVal ws As Worksheet) _
    As Long
'
'==============================================================================
'                                 UsedLastRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the last row of a sheet's used range: the upper bound for every
'   data scan.
'
' INPUTS
'   ws: the worksheet to measure.
'
' RETURNS
'   Row number of the bottom edge of ws.UsedRange.
'
' UPDATED
'   2026-10-06
'==============================================================================
'

'------------------------------------------------------------------------------
' MEASURE
'------------------------------------------------------------------------------
        With ws.UsedRange
            UsedLastRow = .Row + .Rows.Count - 1
        End With

End Function


Private Function ColumnValues( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal lastRow As Long, _
    ByVal colIdx As Long) _
    As Variant
'
'==============================================================================
'                                 ColumnValues
'------------------------------------------------------------------------------
' PURPOSE
'   Read one column range in a single call, always as a 2-D array.
'
' INPUTS
'   ws: the worksheet to read.
'   firstRow, lastRow: rows to read, firstRow <= lastRow.
'   colIdx: table column number.
'
' RETURNS
'   A 1-based array (1 To rows, 1 To 1). Element (r, 1) holds the value of
'   row firstRow + r - 1.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim v     As Variant      'Value or array returned by Range.Value
    Dim a()   As Variant      'One-cell array built for a single-row range

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Range.Value returns a scalar, not an array, for a single cell. Wrap it
    'so callers can always index v(r, 1).
        v = TableRange(ws, firstRow, colIdx, lastRow, colIdx).Value
        If IsArray(v) Then
            ColumnValues = v
        Else
            ReDim a(1 To 1, 1 To 1)
            a(1, 1) = v
            ColumnValues = a
        End If

End Function


Public Function LastDataRow( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal colIdx As Long) _
    As Long
'
'==============================================================================
'                                 LastDataRow
'------------------------------------------------------------------------------
' PURPOSE
'   Find the last non-blank row in one column.
'
' INPUTS
'   ws: the worksheet to scan.
'   firstRow: first row of the data area.
'   colIdx: table column number to test.
'
' RETURNS
'   Row number of the last non-blank cell; firstRow - 1 when the column is
'   empty from firstRow down.
'
' DEPENDENCIES
'   Scans the used range bottom-up instead of using End(xlUp), so that the
'   result is the same in hidden or automated sessions.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim lastR   As Long       'Last row of the used range
    Dim v       As Variant    'Column values from firstRow to lastR
    Dim r       As Long       'Position in v, counted from the bottom

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
    'Start from the "empty" answer, then walk up from the bottom of the used
    'range and stop at the first non-blank cell.
        LastDataRow = firstRow - 1
        lastR = UsedLastRow(ws)
        If lastR < firstRow Then
            Exit Function
        End If
        v = ColumnValues(ws, firstRow, lastR, colIdx)
        For r = UBound(v, 1) To 1 Step -1
            If Not IsBlankCell(v(r, 1)) Then
                LastDataRow = firstRow + r - 1
                Exit Function
            End If
        Next r

End Function


Public Function LastDataRowAny( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal nCols As Long, _
    ByVal ignoreCol As Long) _
    As Long
'
'==============================================================================
'                                LastDataRowAny
'------------------------------------------------------------------------------
' PURPOSE
'   Find the last row with a value in any of the first nCols columns, so
'   that a row whose ID is blank is still found (#35).
'
' INPUTS
'   ws: the worksheet to scan.
'   firstRow: first row of the data area.
'   nCols: number of columns from the first table column.
'   ignoreCol: a column that does not count as data, such as a free-text
'      comment; 0 for none.
'
' RETURNS
'   Row number of the last such row; firstRow - 1 when there is none.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim lastR   As Long       'Last row of the used range
    Dim v       As Variant    'Block values from firstRow to lastR
    Dim r       As Long       'Row in v, counted from the bottom
    Dim c       As Long       'Column in v

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
        LastDataRowAny = firstRow - 1
        lastR = UsedLastRow(ws)
        If lastR < firstRow Then
            Exit Function
        End If
        v = TableRange(ws, firstRow, 1, lastR, nCols).Value
        If Not IsArray(v) Then
            If Not IsBlankCell(v) Then
                LastDataRowAny = firstRow
            End If
            Exit Function
        End If
        For r = UBound(v, 1) To 1 Step -1
            For c = 1 To UBound(v, 2)
                If c <> ignoreCol Then
                    If Not IsBlankCell(v(r, c)) Then
                        LastDataRowAny = firstRow + r - 1
                        Exit Function
                    End If
                End If
            Next c
        Next r

End Function


Public Function RowHasData( _
    ByRef data As Variant, _
    ByVal r As Long, _
    ByVal nCols As Long, _
    ByVal ignoreCol As Long) _
    As Boolean
'
'==============================================================================
'                                  RowHasData
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a row of an input block holds any value.
'
' INPUTS
'   data: a 2-D input block; r: the row in it; nCols: columns to check;
'   ignoreCol: a column that does not count as data; 0 for none.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Long    'Column

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For c = 1 To nCols
            If c <> ignoreCol Then
                If Not IsBlankCell(data(r, c)) Then
                    RowHasData = True
                    Exit Function
                End If
            End If
        Next c

End Function


Public Function FindHeaderRow( _
    ByVal ws As Worksheet, _
    ByVal colIdx As Long, _
    ByVal header As String) _
    As Long
'
'==============================================================================
'                                FindHeaderRow
'------------------------------------------------------------------------------
' PURPOSE
'   Find the row whose cell in one column holds a given text, for example a
'   table header or a parameter code on Params.
'
' INPUTS
'   ws: the worksheet to search.
'   colIdx: table column number to search.
'   header: text to find; compared case-insensitively after trimming.
'
' RETURNS
'   Row number of the first match from the top; 0 when not found.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim lastR   As Long       'Last row of the used range
    Dim v       As Variant    'Column values from row 1 to lastR
    Dim r       As Long       'Row being compared

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        FindHeaderRow = 0
        lastR = UsedLastRow(ws)
        If lastR < 1 Then
            Exit Function
        End If
        v = ColumnValues(ws, 1, lastR, colIdx)
        For r = 1 To UBound(v, 1)
            If StrComp(SafeStr(v(r, 1)), header, vbTextCompare) = 0 Then
                FindHeaderRow = r
                Exit Function
            End If
        Next r

End Function


Public Function GetParam( _
    ByVal code As String) _
    As Variant
'
'==============================================================================
'                                   GetParam
'------------------------------------------------------------------------------
' PURPOSE
'   Read one parameter value through the workbook name equal to its code.
'
' INPUTS
'   code: a PRM_ code from CORE_Config, for example "Alpha".
'
' RETURNS
'   The value of the name's top-left cell; Empty when the name is missing
'   or broken, which CORE_Engine.ValidateSchema reports before any
'   parameter is read.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Range    'The name's top-left cell

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set c = NamedCell(code)
        If Not c Is Nothing Then
            GetParam = c.Value
        End If

End Function


Public Function TextHash( _
    ByVal content As String) _
    As String
'
'==============================================================================
'                                   TextHash
'------------------------------------------------------------------------------
' PURPOSE
'   Reduce a text to an 8-digit hexadecimal fingerprint, to tell whether
'   the inputs of a run have changed since (#36). Not a security hash: two
'   different texts can share a fingerprint, but an edit is missed only
'   about once in two billion.
'
' METHOD
'   Polynomial hash over the UTF-16 bytes of the text, modulo the prime
'   2^31 - 1, computed in Double, which holds every intermediate value
'   exactly.
'
' UPDATED
'   2026-10-07
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const HASH_MODULUS As Double = 2147483647#    'Prime 2^31 - 1
    Const HASH_BASE    As Double = 131#           'Multiplier per byte
    Dim bytes()   As Byte      'UTF-16 bytes of the text
    Dim h         As Double    'Running hash, 0 to HASH_MODULUS - 1
    Dim i         As Long      'Byte being added

'------------------------------------------------------------------------------
' HASH
'------------------------------------------------------------------------------
        bytes = content
        For i = 0 To UBound(bytes)
            h = h * HASH_BASE + bytes(i)
            h = h - Int(h / HASH_MODULUS) * HASH_MODULUS
        Next i
        TextHash = Right$("0000000" & Hex$(CLng(h)), 8)

End Function


Public Sub ClearOutputBlock( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal nCols As Long)
'
'==============================================================================
'                               ClearOutputBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Clear an output table before it is rewritten: values, bold and fill.
'
' INPUTS
'   ws: the output worksheet.
'   firstRow: first data row to clear.
'   nCols: number of columns from the first table column.
'
' STATE OWNERSHIP
'   Changes the cleared cells only. Number formats are left in place.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim lastR   As Long    'Last row of the used range

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        lastR = UsedLastRow(ws)
        If lastR < firstRow Then
            Exit Sub
        End If
        With TableRange(ws, firstRow, 1, lastR, nCols)
            .ClearContents
            .Font.Bold = False
            .Interior.ColorIndex = xlNone
        End With

End Sub


Public Sub WriteBlock( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByRef arr() As Variant, _
    ByVal nRows As Long, _
    ByVal nCols As Long)
'
'==============================================================================
'                                  WriteBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Write the first rows of a 2-D array to a sheet in one call.
'
' INPUTS
'   ws: the output worksheet.
'   firstRow: row receiving arr(1, *); the block starts in the first table
'   column.
'   arr: 1-based 2-D array; it may have more rows than nRows.
'   nRows: number of rows to write; nothing is written when it is 0.
'   nCols: number of columns to write.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Assigning the array to a smaller range writes only its top-left part.
        If nRows <= 0 Then
            Exit Sub
        End If
        TableRange(ws, firstRow, 1, firstRow + nRows - 1, nCols).Value = arr

End Sub


Public Sub FormatColumns( _
    ByVal ws As Worksheet, _
    ByVal firstRow As Long, _
    ByVal nRows As Long, _
    ByVal formats As String)
'
'==============================================================================
'                                FormatColumns
'------------------------------------------------------------------------------
' PURPOSE
'   Apply a number format to each column of an output block.
'
' INPUTS
'   ws: the output worksheet.
'   firstRow: first row of the block.
'   nRows: number of rows; nothing is formatted when it is 0.
'   formats: one number format per table column from the first, separated
'            by "|".
'            An empty entry leaves that column unchanged.
'
' UPDATED
'   2026-10-11
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim parts()   As String    'Number formats, one per column, 0-based
    Dim c         As Long      'Position in parts; column c + 1

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
        If nRows <= 0 Then
            Exit Sub
        End If
        parts = Split(formats, "|")
        For c = 0 To UBound(parts)
            If Len(parts(c)) > 0 Then
                TableRange(ws, firstRow, c + 1, firstRow + nRows - 1, c + 1).NumberFormat = parts(c)
            End If
        Next c

End Sub
