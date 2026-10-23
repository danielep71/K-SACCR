Attribute VB_Name = "CORE_CsvReader"
'==============================================================================
' MODULE: CORE_CsvReader
'------------------------------------------------------------------------------
' PURPOSE
'   Read a Trades or NettingSets CSV file under the CSV input contract
'   (docs/csv/CSV_CONTRACT.md, #206) without Excel converting anything
'   (#207): decode its bytes as UTF-8, split it into records, check the
'   header, every field and the rules between fields, and report every
'   finding with its code, physical line, record and column. A file is
'   accepted or rejected whole; the records of an accepted file are
'   returned as their raw text, in contract column order.
'
' PUBLIC SURFACE
'   None outside this VBA project. ReadCsvFile reads a file;
'   ReadCsvBytes reads bytes already in memory. Both fill a
'   CsvReadResult and the arrays of values, lines and findings it counts.
'   FindingText formats one finding for a message or a test.
'
' DEPENDENCIES
'   CORE_CsvSchema for the contract, generated from docs/csv/contract.json;
'   CORE_Config for the error number. tools/csv_contract.py is the
'   reference implementation: for every fixture and case in tests/csv this
'   module must report exactly its findings (TEST_CsvReader).
'
' STATE OWNERSHIP
'   None. Nothing is written to a sheet, a file or module state.
'
' ERROR POLICY
'   A file that does not conform to the contract is not an error: it gives
'   findings. A file that cannot be found or read, and an unknown kind,
'   raise ERR_CSV_READ.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults. The result does not
'   depend on the Windows locale: the bytes are decoded here, numbers are
'   checked as text and converted with Val, which always reads a "."
'   decimal point, and dates are checked digit by digit. CDbl, IsNumeric,
'   IsDate and Excel's own CSV import are never used.
'
' UPDATED
'   2026-10-23
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the reader off the external API.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Finding codes (docs/csv/CSV_CONTRACT.md#-findings).
        Private Const CSV_ENCODING          As String = "CSV-ENCODING"
        Private Const CSV_SYNTAX            As String = "CSV-SYNTAX"
        Private Const CSV_EMPTY             As String = "CSV-EMPTY"
        Private Const CSV_HEADER_DUPLICATE  As String = "CSV-HEADER-DUPLICATE"
        Private Const CSV_HEADER_MISSING    As String = "CSV-HEADER-MISSING"
        Private Const CSV_HEADER_UNKNOWN    As String = "CSV-HEADER-UNKNOWN"
        Private Const CSV_FIELD_COUNT       As String = "CSV-FIELD-COUNT"
        Private Const CSV_BLANK_RECORD      As String = "CSV-BLANK-RECORD"
        Private Const CSV_TOO_MANY_RECORDS  As String = "CSV-TOO-MANY-RECORDS"
        Private Const CSV_TOO_LARGE         As String = "CSV-TOO-LARGE"
        Private Const CSV_REQUIRED          As String = "CSV-REQUIRED"
        Private Const CSV_TYPE              As String = "CSV-TYPE"
        Private Const CSV_RANGE             As String = "CSV-RANGE"
        Private Const CSV_ALLOWED           As String = "CSV-ALLOWED"
        Private Const CSV_LENGTH            As String = "CSV-LENGTH"
        Private Const CSV_FORMULA           As String = "CSV-FORMULA"
        Private Const CSV_DUPLICATE_KEY     As String = "CSV-DUPLICATE-KEY"
        Private Const CSV_RESERVED_KEY      As String = "CSV-RESERVED-KEY"
        Private Const CSV_ROW_RULE          As String = "CSV-ROW-RULE"

    'Characters a value of a type other than text may not start or end
    'with: those Python's str.strip removes, as the reference does.
        Private Const SPACE_CODES   As String = "|9|10|11|12|13|28|29|30|31|32|133|160|5760|8192|8193|8194|8195|8196|8197|8198|8199|8200|8201|8202|8232|8233|8239|8287|12288|"

    'The digits of the smallest magnitude a cell holds, 2.2250738585072014e-308,
    'cut to CSV_MAX_DIGITS: a value of exponent -308 whose digits are not
    'above them is smaller.
        Private Const SMALLEST_DIGITS    As String = "222507385850720"
        Private Const SMALLEST_EXPONENT  As Long = -308
        Private Const LARGEST_EXPONENT   As Long = 308

'------------------------------------------------------------------------------
' TYPES
'------------------------------------------------------------------------------
    'One finding; record 0 is the header.
        Public Type CsvFinding
            Code            As String     'Stable finding code
            LineNumber      As Long       'Physical line the record starts on; 0 for the whole file
            RecordNumber    As Long       'Record number, the header being 0
            ColumnName      As String     'CSV column name; "" when none applies
            Message         As String     'What is wrong, for the user
        End Type

    'What a read found. The records and findings are separate arrays:
    'Values(1 To RecordCount, 1 To ColumnCount), the raw text in contract
    'order, and Lines(1 To RecordCount), the physical line of each record,
    'both filled only when the file is accepted; Findings(1 To
    'FindingCount), sorted by record, line, column and code.
        Public Type CsvReadResult
            Kind            As String         'Trades or NettingSets
            Accepted        As Boolean        'No finding: every record can be used
            RecordCount     As Long           'Records of an accepted file; 0 otherwise
            ColumnCount     As Long           'Columns of the contract
            FindingCount    As Long           'Number of findings
        End Type


'
'------------------------------------------------------------------------------
'
'                                  READING
'
'------------------------------------------------------------------------------
'

Public Sub ReadCsvFile( _
    ByVal FilePath As String, _
    ByVal Kind As String, _
    ByRef Result As CsvReadResult, _
    ByRef Values() As String, _
    ByRef Lines() As Long, _
    ByRef Findings() As CsvFinding, _
    Optional ByVal MaxBytes As Long = 0)
'
'==============================================================================
'                                 ReadCsvFile
'------------------------------------------------------------------------------
' PURPOSE
'   Read the CSV file FilePath as a file of Kind. A file larger than the
'   contract allows is rejected from its size alone, before any of it is
'   read.
'
' INPUTS
'   FilePath: an existing local file. Kind: "Trades" or "NettingSets".
'   MaxBytes replaces the contract's size limit when above 0, for tests.
'
' RETURNS
'   Result, with Values, Lines and Findings: see CsvReadResult.
'
' ERROR POLICY
'   Raises ERR_CSV_READ for an unknown kind and for a file that does not
'   exist or cannot be read whole, a malformed path included; the file
'   number is closed first.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Spec            As CsvSpec      'The file's contract
    Dim Columns()       As CsvColumn    'Its columns
    Dim Rules()         As CsvRule      'Its rules
    Dim ByteCount       As Long         'Size of the file, in bytes
    Dim Bytes()         As Byte         'Its content
    Dim FileNumber      As Integer      'File number; 0 while none is open
    Dim ErrorNumber     As Long         'First error
    Dim ErrorText       As String       'Its description

'------------------------------------------------------------------------------
' CHECK THE KIND AND THE SIZE
'------------------------------------------------------------------------------
        If Not CsvSpecOf(Kind, Spec, Columns, Rules) Then
            Err.Raise ERR_CSV_READ, "CORE_CsvReader.ReadCsvFile", "unknown CSV kind '" & Kind & "'"
        End If
        If Len(FilePath) = 0 Then
            Err.Raise ERR_CSV_READ, "CORE_CsvReader.ReadCsvFile", "no file given"
        End If
        If MaxBytes <= 0 Then
            MaxBytes = CSV_MAX_BYTES
        End If
    'Dir$ itself raises for a malformed or inaccessible path, so the handler
    'is in place before it (#446).
        On Error GoTo Failed
        If Len(Dir$(FilePath, vbHidden Or vbSystem Or vbReadOnly)) = 0 Then
            Err.Raise ERR_CSV_READ, "CORE_CsvReader.ReadCsvFile", "the file does not exist: " & FilePath
        End If
        ByteCount = FileLen(FilePath)
        If ByteCount < 0 Or ByteCount > MaxBytes Then
            On Error GoTo 0
            StartResult Result, Spec, Values, Lines, Findings
            AddFinding Result, Findings, CSV_TOO_LARGE, 0, 0, "", "the file has more than " & MaxBytes & " bytes"
            Exit Sub
        End If

'------------------------------------------------------------------------------
' READ THE FILE WHOLE
'------------------------------------------------------------------------------
        If ByteCount > 0 Then
            ReDim Bytes(0 To ByteCount - 1)
            FileNumber = FreeFile
            Open FilePath For Binary Access Read Shared As #FileNumber
            If LOF(FileNumber) <> ByteCount Then
                Err.Raise ERR_CSV_READ, "CORE_CsvReader.ReadCsvFile", "the file changed while it was read"
            End If
            Get #FileNumber, 1, Bytes
            Close #FileNumber
            FileNumber = 0
        End If
        On Error GoTo 0
        ReadCsvBytes Bytes, ByteCount, Kind, Result, Values, Lines, Findings, 0, MaxBytes
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        ErrorNumber = Err.Number
        ErrorText = Err.Description
        If FileNumber <> 0 Then
            Close #FileNumber
        End If
        On Error GoTo 0
        If ErrorNumber = ERR_CSV_READ Then
            Err.Raise ERR_CSV_READ, "CORE_CsvReader.ReadCsvFile", ErrorText
        End If
        Err.Raise ERR_CSV_READ, "CORE_CsvReader.ReadCsvFile", _
                  "Could not read " & FilePath & ": " & ErrorText & " (error " & ErrorNumber & ")"

End Sub


Public Sub ReadCsvBytes( _
    ByRef Bytes() As Byte, _
    ByVal ByteCount As Long, _
    ByVal Kind As String, _
    ByRef Result As CsvReadResult, _
    ByRef Values() As String, _
    ByRef Lines() As Long, _
    ByRef Findings() As CsvFinding, _
    Optional ByVal MaxRecords As Long = 0, _
    Optional ByVal MaxBytes As Long = 0)
'
'==============================================================================
'                                ReadCsvBytes
'------------------------------------------------------------------------------
' PURPOSE
'   Read the first ByteCount bytes of Bytes as a CSV file of Kind, as
'   tools/csv_contract.py validates it: size, encoding, records, header,
'   record count, fields, rules between fields and keys, in that order.
'   A header error, a record limit exceeded, a file too large or not valid
'   UTF-8 stops reading; every other finding is reported.
'
' INPUTS
'   Bytes: the file's content, from index 0; unallocated when ByteCount
'   is 0. ByteCount >= 0. Kind: "Trades" or "NettingSets". MaxRecords and
'   MaxBytes replace the contract's limits when above 0, for tests.
'
' RETURNS
'   Result, with Values, Lines and Findings: see CsvReadResult.
'
' ERROR POLICY
'   Raises ERR_CSV_READ for an unknown kind.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Spec            As CsvSpec      'The file's contract
    Dim Columns()       As CsvColumn    'Its columns
    Dim Rules()         As CsvRule      'Its rules
    Dim Text            As String       'The decoded text
    Dim NulAt           As Long         'Position of a NUL character; 0 for none
    Dim RecordFields()  As Variant      'Fields of each record, as String arrays; 1 is the header
    Dim RecordLines()   As Long         'Physical line each record starts on
    Dim RecordTotal     As Long         'Records parsed, the header included
    Dim SyntaxLine      As Long         'Line of the record with a syntax error; 0 for none
    Dim SyntaxMessage   As String       'What the syntax error is
    Dim Limit           As Long         'Records allowed after the header

'------------------------------------------------------------------------------
' SIZE AND ENCODING
'------------------------------------------------------------------------------
        If Not CsvSpecOf(Kind, Spec, Columns, Rules) Then
            Err.Raise ERR_CSV_READ, "CORE_CsvReader.ReadCsvBytes", "unknown CSV kind '" & Kind & "'"
        End If
        StartResult Result, Spec, Values, Lines, Findings
        If MaxBytes <= 0 Then
            MaxBytes = CSV_MAX_BYTES
        End If
        If ByteCount > MaxBytes Then
            AddFinding Result, Findings, CSV_TOO_LARGE, 0, 0, "", ByteCount & " bytes, at most " & MaxBytes
            Exit Sub
        End If
        If Not DecodeUtf8(Bytes, ByteCount, Text) Then
            AddFinding Result, Findings, CSV_ENCODING, 0, 0, "", "not valid UTF-8"
            Exit Sub
        End If
        NulAt = InStr(1, Text, vbNullChar, vbBinaryCompare)
        If NulAt > 0 Then
            AddFinding Result, Findings, CSV_ENCODING, CountLineFeeds(Text, NulAt - 1) + 1, 0, "", _
                       "contains a NUL character"
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RECORDS AND HEADER
'------------------------------------------------------------------------------
    'The header and one record past the limit are all a rejected file needs:
    'the parser counts the others without keeping them.
        Limit = IIf(MaxRecords > 0, MaxRecords, Spec.MaxRecords)
        ParseRecords Text, Limit + 2, RecordFields, RecordLines, RecordTotal, SyntaxLine, SyntaxMessage
        If RecordTotal = 0 Then
            If SyntaxLine > 0 Then
                AddFinding Result, Findings, CSV_SYNTAX, SyntaxLine, 0, "", SyntaxMessage
            Else
                AddFinding Result, Findings, CSV_EMPTY, 1, 0, "", "the file has no header record"
            End If
            Exit Sub
        End If
        If Not HeaderAccepted(RecordFields(1), RecordLines(1), Spec, Columns, Result, Findings) Then
            Exit Sub
        End If
        If SyntaxLine > 0 Then
            AddFinding Result, Findings, CSV_SYNTAX, SyntaxLine, RecordTotal, "", SyntaxMessage
        End If
        If RecordTotal - 1 > Limit Then
            AddFinding Result, Findings, CSV_TOO_MANY_RECORDS, RecordLines(Limit + 2), Limit + 1, "", _
                       (RecordTotal - 1) & " records, at most " & Limit
            SortFindings Result, Findings
            Exit Sub
        End If

'------------------------------------------------------------------------------
' FIELDS, RULES AND KEYS
'------------------------------------------------------------------------------
        CheckBody RecordFields, RecordLines, RecordTotal, Spec, Columns, Rules, Result, Values, Lines, Findings
        SortFindings Result, Findings
        Result.Accepted = (Result.FindingCount = 0)
        If Not Result.Accepted Then
            Result.RecordCount = 0
            Erase Values
            Erase Lines
        End If

End Sub


Public Function FindingText( _
    ByRef Finding As CsvFinding) _
    As String
'
'==============================================================================
'                                 FindingText
'------------------------------------------------------------------------------
' PURPOSE
'   One finding as "CODE|line|record|column", the form the tests compare.
'
' UPDATED
'   2026-10-23
'==============================================================================
'
        FindingText = Finding.Code & "|" & Finding.LineNumber & "|" & Finding.RecordNumber & "|" & Finding.ColumnName

End Function


'
'------------------------------------------------------------------------------
'
'                                  DECODING
'
'------------------------------------------------------------------------------
'

Private Function DecodeUtf8( _
    ByRef Bytes() As Byte, _
    ByVal ByteCount As Long, _
    ByRef Text As String) _
    As Boolean
'
'==============================================================================
'                                 DecodeUtf8
'------------------------------------------------------------------------------
' PURPOSE
'   Decode the first ByteCount bytes as strict UTF-8, as Python's
'   bytes.decode("utf-8") does, after skipping a byte-order mark: an
'   overlong form, a surrogate, a code point above U+10FFFF, a stray or
'   missing continuation byte and a truncated sequence are all invalid.
'   Code points above U+FFFF become surrogate pairs.
'
' RETURNS
'   True and Text when the bytes are valid; False otherwise.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Units()         As Byte       'UTF-16LE code units being written
    Dim UnitBytes       As Long       'Bytes of Units written
    Dim i               As Long       'Position in Bytes
    Dim Lead            As Long       'Lead byte of a sequence
    Dim Needed          As Long       'Continuation bytes the lead byte announces
    Dim Lowest          As Long       'Smallest valid second byte
    Dim Highest         As Long       'Largest valid second byte
    Dim CodePoint       As Long       'Code point decoded
    Dim k               As Long       'Continuation byte counter
    Dim Follower        As Long       'A continuation byte

'------------------------------------------------------------------------------
' DECODE
'------------------------------------------------------------------------------
        Text = vbNullString
        If ByteCount >= 3 Then
            If Bytes(0) = &HEF And Bytes(1) = &HBB And Bytes(2) = &HBF Then
                i = 3
            End If
        End If
        If i >= ByteCount Then
            DecodeUtf8 = True
            Exit Function
        End If
        ReDim Units(0 To 2 * (ByteCount - i) - 1)
        Do While i < ByteCount
            Lead = Bytes(i)
            Lowest = &H80
            Highest = &HBF
            Select Case Lead
                Case Is < &H80
                    Needed = 0
                    CodePoint = Lead
                Case &HC2 To &HDF
                    Needed = 1
                    CodePoint = Lead And &H1F
                Case &HE0 To &HEF
                    Needed = 2
                    CodePoint = Lead And &HF
                    If Lead = &HE0 Then Lowest = &HA0
                    If Lead = &HED Then Highest = &H9F
                Case &HF0 To &HF4
                    Needed = 3
                    CodePoint = Lead And &H7
                    If Lead = &HF0 Then Lowest = &H90
                    If Lead = &HF4 Then Highest = &H8F
                Case Else
                    Exit Function
            End Select
            If i + Needed >= ByteCount And Needed > 0 Then
                Exit Function
            End If
            For k = 1 To Needed
                Follower = Bytes(i + k)
                If k = 1 Then
                    If Follower < Lowest Or Follower > Highest Then Exit Function
                ElseIf Follower < &H80 Or Follower > &HBF Then
                    Exit Function
                End If
                CodePoint = CodePoint * 64 + (Follower And &H3F)
            Next k
            i = i + Needed + 1
            If CodePoint >= &H10000 Then
                CodePoint = CodePoint - &H10000
                WriteUnit Units, UnitBytes, &HD800& + CodePoint \ 1024
                WriteUnit Units, UnitBytes, &HDC00& + (CodePoint And &H3FF)
            Else
                WriteUnit Units, UnitBytes, CodePoint
            End If
        Loop
        ReDim Preserve Units(0 To UnitBytes - 1)
        Text = Units
        DecodeUtf8 = True

End Function


Private Sub WriteUnit( _
    ByRef Units() As Byte, _
    ByRef UnitBytes As Long, _
    ByVal CodeUnit As Long)
'
'==============================================================================
' PURPOSE
'   Append one UTF-16 code unit, low byte first.
'==============================================================================
'
        Units(UnitBytes) = CodeUnit And &HFF
        Units(UnitBytes + 1) = CodeUnit \ 256
        UnitBytes = UnitBytes + 2

End Sub


Private Function CountLineFeeds( _
    ByRef Text As String, _
    ByVal LastPosition As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The number of line feeds in the first LastPosition characters of Text.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Found   As Long    'Position of the next line feed

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        Found = InStr(1, Text, vbLf, vbBinaryCompare)
        Do While Found > 0 And Found <= LastPosition
            CountLineFeeds = CountLineFeeds + 1
            Found = InStr(Found + 1, Text, vbLf, vbBinaryCompare)
        Loop

End Function


'
'------------------------------------------------------------------------------
'
'                                  PARSING
'
'------------------------------------------------------------------------------
'

Private Sub ParseRecords( _
    ByRef Text As String, _
    ByVal MaxKept As Long, _
    ByRef RecordFields() As Variant, _
    ByRef RecordLines() As Long, _
    ByRef RecordTotal As Long, _
    ByRef SyntaxLine As Long, _
    ByRef SyntaxMessage As String)
'
'==============================================================================
'                                ParseRecords
'------------------------------------------------------------------------------
' PURPOSE
'   Split Text into records of fields: ";" between fields, CRLF or LF
'   between records, double quotes around a whole field, two of them for
'   one inside it. Parsing stops at the first syntax error: a quote inside
'   an unquoted field, text after a closing quote, a carriage return that
'   does not end a line, or a quoted field still open at the end.
'
' INPUTS
'   MaxKept: the records kept, the header included; later records are
'   parsed and counted but not kept, so that a file of many short records
'   cannot exhaust memory before its record limit is reported.
'
' RETURNS
'   RecordTotal, the records parsed; RecordFields(1 To the smaller of
'   RecordTotal and MaxKept), each a String array from 0, and the physical
'   line each kept record starts on; SyntaxLine and SyntaxMessage for a
'   syntax error, SyntaxLine 0 when there is none.
'
' METHOD
'   Each delimiter's next position is found with InStr and kept until the
'   parser passes it, so the text is scanned once per delimiter however
'   far apart they are. A quoted field is taken whole once its closing
'   quote is found, its doubled quotes replaced in one pass.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TextLength      As Long         'Characters of Text
    Dim Position        As Long         'Next character to read
    Dim CurrentLine     As Long         'Physical line of Position
    Dim StartLine       As Long         'Line the record starts on
    Dim QuoteLine       As Long         'Line a quoted field opens on
    Dim NextQuote       As Long         'Next double quote; 0 to look again, -1 for none
    Dim NextSemicolon   As Long         'Next ";"
    Dim NextLineFeed    As Long         'Next line feed
    Dim NextReturn      As Long         'Next carriage return
    Dim CloseAt         As Long         'Closing quote of a quoted field
    Dim FieldStart      As Long         'First character inside a quoted field
    Dim StopAt          As Long         'End of an unquoted field
    Dim Value           As String       'The field being read
    Dim Fields()        As String       'Fields of the record being read
    Dim FieldCount      As Long         'Fields read
    Dim Following       As String       'Character after a closing quote
    Dim Ending          As Long         'Length of the line end at Position

'------------------------------------------------------------------------------
' READ RECORD BY RECORD
'------------------------------------------------------------------------------
        TextLength = Len(Text)
        Position = 1
        CurrentLine = 1
        RecordTotal = 0
        ReDim RecordFields(1 To 16)
        ReDim RecordLines(1 To 16)
        ReDim Fields(0 To 31)
        Do While Position <= TextLength
            StartLine = CurrentLine
            FieldCount = 0
            Do
                If Mid$(Text, Position, 1) = """" Then
                    'A quoted field, up to the quote that is not doubled.
                    QuoteLine = CurrentLine
                    Position = Position + 1
                    FieldStart = Position
                    Do
                        CloseAt = NextOf(Text, """", Position, NextQuote)
                        If CloseAt = 0 Then
                            SyntaxLine = StartLine
                            SyntaxMessage = "quoted field opened on line " & QuoteLine & " is not closed"
                            Exit Sub
                        End If
                        CurrentLine = CurrentLine + LineFeedsBetween(Text, Position, CloseAt - 1, NextLineFeed)
                        If Mid$(Text, CloseAt + 1, 1) = """" Then
                            Position = CloseAt + 2
                        Else
                            Position = CloseAt + 1
                            Exit Do
                        End If
                    Loop
                    Value = Replace(Mid$(Text, FieldStart, CloseAt - FieldStart), """""", """", 1, -1, vbBinaryCompare)
                    Following = Mid$(Text, Position, 1)
                    If Len(Following) > 0 And Following <> CSV_DELIMITER And LineEndLength(Text, Position) = 0 Then
                        SyntaxLine = StartLine
                        SyntaxMessage = "text after a closing quote"
                        Exit Sub
                    End If
                Else
                    'An unquoted field, up to the first delimiter, line end or quote.
                    StopAt = TextLength + 1
                    StopAt = Earliest(StopAt, NextOf(Text, CSV_DELIMITER, Position, NextSemicolon))
                    StopAt = Earliest(StopAt, NextOf(Text, vbLf, Position, NextLineFeed))
                    StopAt = Earliest(StopAt, NextOf(Text, vbCr, Position, NextReturn))
                    StopAt = Earliest(StopAt, NextOf(Text, """", Position, NextQuote))
                    If StopAt <= TextLength Then
                        If Mid$(Text, StopAt, 1) = """" Then
                            SyntaxLine = StartLine
                            SyntaxMessage = "quote inside an unquoted field"
                            Exit Sub
                        End If
                        If Mid$(Text, StopAt, 1) = vbCr And Mid$(Text, StopAt + 1, 1) <> vbLf Then
                            SyntaxLine = StartLine
                            SyntaxMessage = "carriage return that does not end a line"
                            Exit Sub
                        End If
                    End If
                    Value = Mid$(Text, Position, StopAt - Position)
                    Position = StopAt
                End If
                If FieldCount > UBound(Fields) Then
                    ReDim Preserve Fields(0 To 2 * FieldCount - 1)
                End If
                Fields(FieldCount) = Value
                FieldCount = FieldCount + 1
                If Mid$(Text, Position, 1) = CSV_DELIMITER Then
                    Position = Position + 1
                Else
                    Exit Do
                End If
            Loop
            Ending = LineEndLength(Text, Position)
            Position = Position + Ending
            If Ending > 0 Then
                CurrentLine = CurrentLine + 1
            End If
            AddRecord RecordFields, RecordLines, RecordTotal, MaxKept, Fields, FieldCount, StartLine
        Loop

End Sub


Private Function NextOf( _
    ByRef Text As String, _
    ByVal Character As String, _
    ByVal Position As Long, _
    ByRef Cached As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The position of the next Character at or after Position, 0 when there
'   is none. Cached keeps the last position found: 0 to look again, -1 when
'   the text holds no more.
'==============================================================================
'
        If Cached = -1 Then
            Exit Function
        End If
        If Cached < Position Then
            Cached = InStr(Position, Text, Character, vbBinaryCompare)
            If Cached = 0 Then
                Cached = -1
                Exit Function
            End If
        End If
        NextOf = Cached

End Function


Private Function LineFeedsBetween( _
    ByRef Text As String, _
    ByVal FirstPosition As Long, _
    ByVal LastPosition As Long, _
    ByRef Cached As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The number of line feeds from FirstPosition to LastPosition, through
'   the line-feed position kept by the parser.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Found   As Long    'Position of the next line feed

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        Found = NextOf(Text, vbLf, FirstPosition, Cached)
        Do While Found > 0 And Found <= LastPosition
            LineFeedsBetween = LineFeedsBetween + 1
            Found = NextOf(Text, vbLf, Found + 1, Cached)
        Loop

End Function


Private Function Earliest( _
    ByVal Current As Long, _
    ByVal Candidate As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The smaller of two positions, a Candidate of 0 meaning none.
'==============================================================================
'
        Earliest = Current
        If Candidate > 0 And Candidate < Current Then
            Earliest = Candidate
        End If

End Function


Private Function LineEndLength( _
    ByRef Text As String, _
    ByVal Position As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The length of the line end at Position: 2 for CRLF, 1 for LF, 0 for
'   none.
'==============================================================================
'
        If Mid$(Text, Position, 2) = vbCrLf Then
            LineEndLength = 2
        ElseIf Mid$(Text, Position, 1) = vbLf Then
            LineEndLength = 1
        End If

End Function


Private Sub AddRecord( _
    ByRef RecordFields() As Variant, _
    ByRef RecordLines() As Long, _
    ByRef RecordTotal As Long, _
    ByVal MaxKept As Long, _
    ByRef Fields() As String, _
    ByVal FieldCount As Long, _
    ByVal StartLine As Long)
'
'==============================================================================
' PURPOSE
'   Count a record of FieldCount fields and keep it while fewer than
'   MaxKept are kept, doubling the arrays when full.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Copy()  As String    'The record's fields
    Dim c       As Long      'Field counter

'------------------------------------------------------------------------------
' APPEND
'------------------------------------------------------------------------------
        RecordTotal = RecordTotal + 1
        If RecordTotal > MaxKept Then
            Exit Sub
        End If
        If RecordTotal > UBound(RecordLines) Then
            ReDim Preserve RecordFields(1 To 2 * UBound(RecordLines))
            ReDim Preserve RecordLines(1 To 2 * UBound(RecordLines))
        End If
        ReDim Copy(0 To FieldCount - 1)
        For c = 0 To FieldCount - 1
            Copy(c) = Fields(c)
        Next c
        RecordFields(RecordTotal) = Copy
        RecordLines(RecordTotal) = StartLine

End Sub


'
'------------------------------------------------------------------------------
'
'                                  CHECKING
'
'------------------------------------------------------------------------------
'

Private Function HeaderAccepted( _
    ByRef Header As Variant, _
    ByVal HeaderLine As Long, _
    ByRef Spec As CsvSpec, _
    ByRef Columns() As CsvColumn, _
    ByRef Result As CsvReadResult, _
    ByRef Findings() As CsvFinding) _
    As Boolean
'
'==============================================================================
'                               HeaderAccepted
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the header names every contract column exactly once, in
'   any order, with no other column; names are case-sensitive. Reports a
'   repeated or unknown name in header order, then each missing column in
'   contract order.
'
' RETURNS
'   True when the header has no finding.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Seen()      As Boolean    'Contract columns already named
    Dim h           As Long       'Header field counter
    Dim c           As Long       'Contract column counter
    Dim Earlier     As Long       'Earlier header field counter
    Dim Repeated    As Boolean    'The name was already in the header

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        ReDim Seen(1 To Spec.ColumnCount)
        For h = 0 To UBound(Header)
            Repeated = False
            For Earlier = 0 To h - 1
                If StrComp(Header(Earlier), Header(h), vbBinaryCompare) = 0 Then
                    Repeated = True
                    Exit For
                End If
            Next Earlier
            c = ColumnIndex(Spec, Columns, Header(h))
            If Repeated Then
                AddFinding Result, Findings, CSV_HEADER_DUPLICATE, HeaderLine, 0, Header(h), _
                           "column '" & Header(h) & "' repeats"
            ElseIf c = 0 Then
                AddFinding Result, Findings, CSV_HEADER_UNKNOWN, HeaderLine, 0, Header(h), _
                           "column '" & Header(h) & "' is not in the contract"
            End If
            If c > 0 Then
                Seen(c) = True
            End If
        Next h
        For c = 1 To Spec.ColumnCount
            If Not Seen(c) Then
                AddFinding Result, Findings, CSV_HEADER_MISSING, HeaderLine, 0, Columns(c).ColumnName, _
                           "column '" & Columns(c).ColumnName & "' is missing"
            End If
        Next c
        HeaderAccepted = (Result.FindingCount = 0)

End Function


Private Sub CheckBody( _
    ByRef RecordFields() As Variant, _
    ByRef RecordLines() As Long, _
    ByVal RecordTotal As Long, _
    ByRef Spec As CsvSpec, _
    ByRef Columns() As CsvColumn, _
    ByRef Rules() As CsvRule, _
    ByRef Result As CsvReadResult, _
    ByRef Values() As String, _
    ByRef Lines() As Long, _
    ByRef Findings() As CsvFinding)
'
'==============================================================================
'                                  CheckBody
'------------------------------------------------------------------------------
' PURPOSE
'   Check every record after the header: an empty line, the field count,
'   then each field and the rules between fields (CheckRecord); then the
'   keys across the records. Keeps each record with the header's field
'   count in Values and Lines, in contract column order.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header      As Variant    'Header fields
    Dim Position()  As Long       'Header position of each contract column
    Dim Fields      As Variant    'Fields of a record
    Dim Row()       As String     'One record in contract order
    Dim RowCount    As Long       'Records kept
    Dim RowRecord() As Long       'Record number of each kept record
    Dim r           As Long       'Record counter, the header being 1
    Dim c           As Long       'Contract column counter

'------------------------------------------------------------------------------
' MAP THE HEADER
'------------------------------------------------------------------------------
        Header = RecordFields(1)
        ReDim Position(1 To Spec.ColumnCount)
        For c = 1 To Spec.ColumnCount
            For r = 0 To UBound(Header)
                If StrComp(Header(r), Columns(c).ColumnName, vbBinaryCompare) = 0 Then
                    Position(c) = r
                    Exit For
                End If
            Next r
        Next c
        If RecordTotal > 1 Then
            ReDim Values(1 To RecordTotal - 1, 1 To Spec.ColumnCount)
            ReDim Lines(1 To RecordTotal - 1)
            ReDim RowRecord(1 To RecordTotal - 1)
        End If
        ReDim Row(1 To Spec.ColumnCount)

'------------------------------------------------------------------------------
' CHECK EACH RECORD
'------------------------------------------------------------------------------
        For r = 2 To RecordTotal
            Fields = RecordFields(r)
            If UBound(Fields) = 0 And Len(Fields(0)) = 0 Then
                AddFinding Result, Findings, CSV_BLANK_RECORD, RecordLines(r), r - 1, "", "empty line"
            ElseIf UBound(Fields) <> UBound(Header) Then
                AddFinding Result, Findings, CSV_FIELD_COUNT, RecordLines(r), r - 1, "", _
                           (UBound(Fields) + 1) & " fields, the header has " & (UBound(Header) + 1)
            Else
                For c = 1 To Spec.ColumnCount
                    Row(c) = Fields(Position(c))
                Next c
                CheckRecord Row, RecordLines(r), r - 1, Spec, Columns, Rules, Result, Findings
                RowCount = RowCount + 1
                For c = 1 To Spec.ColumnCount
                    Values(RowCount, c) = Row(c)
                Next c
                Lines(RowCount) = RecordLines(r)
                RowRecord(RowCount) = r - 1
            End If
        Next r
        Result.RecordCount = RowCount
        If RowCount > 0 Then
            CheckKeys Spec, Columns, RowCount, RowRecord, Result, Values, Lines, Findings
        End If

End Sub


Private Sub CheckRecord( _
    ByRef Row() As String, _
    ByVal LineNumber As Long, _
    ByVal RecordNumber As Long, _
    ByRef Spec As CsvSpec, _
    ByRef Columns() As CsvColumn, _
    ByRef Rules() As CsvRule, _
    ByRef Result As CsvReadResult, _
    ByRef Findings() As CsvFinding)
'
'==============================================================================
'                                 CheckRecord
'------------------------------------------------------------------------------
' PURPOSE
'   Check one record in contract order: a record whose fields are all
'   blank is one finding; otherwise each required field must be filled,
'   each filled field must be of its column's type, and each rule between
'   fields must hold. A rule that reads a field with a finding is not
'   judged.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Flawed()    As Boolean    'Columns with a finding
    Dim Blank       As Boolean    'Every field is blank
    Dim Code        As String     'Finding code of a field
    Dim Message     As String     'Its message
    Dim c           As Long       'Contract column counter

'------------------------------------------------------------------------------
' BLANK RECORD
'------------------------------------------------------------------------------
        Blank = True
        For c = 1 To Spec.ColumnCount
            If Len(Row(c)) > 0 Then
                Blank = False
                Exit For
            End If
        Next c
        If Blank Then
            AddFinding Result, Findings, CSV_BLANK_RECORD, LineNumber, RecordNumber, "", "every field is blank"
            Exit Sub
        End If

'------------------------------------------------------------------------------
' FIELDS
'------------------------------------------------------------------------------
        ReDim Flawed(1 To Spec.ColumnCount)
        For c = 1 To Spec.ColumnCount
            If Len(Row(c)) = 0 Then
                If Columns(c).Required Then
                    AddFinding Result, Findings, CSV_REQUIRED, LineNumber, RecordNumber, Columns(c).ColumnName, _
                               "required field is blank"
                    Flawed(c) = True
                End If
            ElseIf FieldProblem(Row(c), Columns(c), Code, Message) Then
                AddFinding Result, Findings, Code, LineNumber, RecordNumber, Columns(c).ColumnName, Message
                Flawed(c) = True
            End If
        Next c

'------------------------------------------------------------------------------
' RULES BETWEEN FIELDS
'------------------------------------------------------------------------------
        For c = 1 To Spec.RuleCount
            If Not RuleReadsFlawed(Spec, Columns, Rules(c), Flawed) Then
                CheckRule Row, Spec, Columns, Rules(c), LineNumber, RecordNumber, Result, Findings
            End If
        Next c

End Sub


Private Function RuleReadsFlawed( _
    ByRef Spec As CsvSpec, _
    ByRef Columns() As CsvColumn, _
    ByRef Rule As CsvRule, _
    ByRef Flawed() As Boolean) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Whether any column the rule reads has a finding.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Names   As Variant    'Columns the rule reads
    Dim n       As Long       'Name counter

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        Names = Split(Rule.Columns, "|")
        For n = 0 To UBound(Names)
            If Flawed(ColumnIndex(Spec, Columns, CStr(Names(n)))) Then
                RuleReadsFlawed = True
                Exit Function
            End If
        Next n

End Function


Private Sub CheckRule( _
    ByRef Row() As String, _
    ByRef Spec As CsvSpec, _
    ByRef Columns() As CsvColumn, _
    ByRef Rule As CsvRule, _
    ByVal LineNumber As Long, _
    ByVal RecordNumber As Long, _
    ByRef Result As CsvReadResult, _
    ByRef Findings() As CsvFinding)
'
'==============================================================================
'                                  CheckRule
'------------------------------------------------------------------------------
' PURPOSE
'   Judge one rule between fields of a record whose fields it reads have
'   no finding. "when": while the condition holds, each Required field must
'   be filled; otherwise each OtherwiseBlank field must be blank. "after"
'   and "above": when both fields are filled, the Later date or number
'   must be after or above the Earlier one.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WhenNames   As Variant    'Condition columns
    Dim WhenValues  As Variant    'Their values
    Dim Names       As Variant    'Columns the rule constrains
    Dim Active      As Boolean    'The condition holds
    Dim LaterText   As String     'Later field
    Dim EarlierText As String     'Earlier field
    Dim Ordered     As Boolean    'Later is after or above Earlier
    Dim n           As Long       'Name counter

'------------------------------------------------------------------------------
' JUDGE
'------------------------------------------------------------------------------
        If Rule.Kind = "when" Then
            WhenNames = Split(Rule.WhenColumns, "|")
            WhenValues = Split(Rule.WhenValues, "|")
            Active = True
            For n = 0 To UBound(WhenNames)
                If StrComp(Row(ColumnIndex(Spec, Columns, CStr(WhenNames(n)))), CStr(WhenValues(n)), vbBinaryCompare) <> 0 Then
                    Active = False
                End If
            Next n
            Names = Split(IIf(Active, Rule.Required, Rule.OtherwiseBlank), "|")
            For n = 0 To UBound(Names)
                If Active And Len(Row(ColumnIndex(Spec, Columns, CStr(Names(n))))) = 0 Then
                    AddFinding Result, Findings, CSV_ROW_RULE, LineNumber, RecordNumber, CStr(Names(n)), _
                               Rule.Id & ": required when " & Rule.WhenColumns & " is " & Rule.WhenValues
                ElseIf Not Active And Len(Row(ColumnIndex(Spec, Columns, CStr(Names(n))))) > 0 Then
                    AddFinding Result, Findings, CSV_ROW_RULE, LineNumber, RecordNumber, CStr(Names(n)), _
                               Rule.Id & ": must be blank unless " & Rule.WhenColumns & " is " & Rule.WhenValues
                End If
            Next n
        Else
            LaterText = Row(ColumnIndex(Spec, Columns, Rule.Later))
            EarlierText = Row(ColumnIndex(Spec, Columns, Rule.Earlier))
            If Len(LaterText) > 0 And Len(EarlierText) > 0 Then
                If Rule.Kind = "after" Then
                    Ordered = (StrComp(LaterText, EarlierText, vbBinaryCompare) > 0)
                Else
                    Ordered = (Val(LaterText) > Val(EarlierText))
                End If
                If Not Ordered Then
                    AddFinding Result, Findings, CSV_ROW_RULE, LineNumber, RecordNumber, Rule.Later, _
                               Rule.Id & ": " & Rule.Later & " must be " & Rule.Kind & " " & Rule.Earlier
                End If
            End If
        End If

End Sub


Private Sub CheckKeys( _
    ByRef Spec As CsvSpec, _
    ByRef Columns() As CsvColumn, _
    ByVal RowCount As Long, _
    ByRef RowRecord() As Long, _
    ByRef Result As CsvReadResult, _
    ByRef Values() As String, _
    ByRef Lines() As Long, _
    ByRef Findings() As CsvFinding)
'
'==============================================================================
'                                  CheckKeys
'------------------------------------------------------------------------------
' PURPOSE
'   Report a reserved key and a key that repeats an earlier record. Keys
'   are compared in upper case, as the engine compares IDs with UCase$;
'   blank keys are left to the required-field check.
'
' METHOD
'   An open-addressing hash table of the upper-case keys, compared as
'   binary strings, so that no locale rule makes two keys equal.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim KeyColumn       As Long       'Contract column of the key
    Dim Reserved        As Variant    'Reserved keys
    Dim TableSize       As Long       'Slots of the hash table, a power of two
    Dim TableKeys()     As String     'Upper-case key in each slot
    Dim TableRecord()   As Long       'Record of the first use; 0 for an empty slot
    Dim Key             As String     'The key, as written
    Dim Folded          As String     'The key in upper case
    Dim Slot            As Long       'Slot probed
    Dim IsReserved      As Boolean    'The key is reserved
    Dim r               As Long       'Kept record counter
    Dim n               As Long       'Reserved key counter

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
        KeyColumn = ColumnIndex(Spec, Columns, Spec.Key)
        Reserved = Split(Spec.ReservedKeys, "|")
        TableSize = 16
        Do While TableSize < 2 * RowCount
            TableSize = TableSize * 2
        Loop
        ReDim TableKeys(0 To TableSize - 1)
        ReDim TableRecord(0 To TableSize - 1)

'------------------------------------------------------------------------------
' CHECK EACH KEY
'------------------------------------------------------------------------------
        For r = 1 To RowCount
            Key = Values(r, KeyColumn)
            If Len(Key) > 0 Then
                Folded = UCase$(Key)
                IsReserved = False
                For n = 0 To UBound(Reserved)
                    If StrComp(Folded, UCase$(CStr(Reserved(n))), vbBinaryCompare) = 0 Then
                        IsReserved = True
                    End If
                Next n
                If IsReserved Then
                    AddFinding Result, Findings, CSV_RESERVED_KEY, Lines(r), RowRecord(r), Spec.Key, _
                               "'" & Key & "' is reserved"
                Else
                    Slot = HashOf(Folded, TableSize)
                    Do While TableRecord(Slot) > 0
                        If StrComp(TableKeys(Slot), Folded, vbBinaryCompare) = 0 Then
                            Exit Do
                        End If
                        Slot = (Slot + 1) And (TableSize - 1)
                    Loop
                    If TableRecord(Slot) > 0 Then
                        AddFinding Result, Findings, CSV_DUPLICATE_KEY, Lines(r), RowRecord(r), Spec.Key, _
                                   "'" & Key & "' repeats record " & TableRecord(Slot)
                    Else
                        TableKeys(Slot) = Folded
                        TableRecord(Slot) = RowRecord(r)
                    End If
                End If
            End If
        Next r

End Sub


Private Function HashOf( _
    ByRef Key As String, _
    ByVal TableSize As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   A slot from 0 to TableSize - 1 for Key; TableSize is a power of two up
'   to 2^24, so that no intermediate value overflows a Long.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i       As Long    'Character counter

'------------------------------------------------------------------------------
' HASH
'------------------------------------------------------------------------------
        For i = 1 To Len(Key)
            HashOf = (HashOf * 31 + (AscW(Mid$(Key, i, 1)) And &HFFFF&)) And (TableSize - 1)
        Next i

End Function


'
'------------------------------------------------------------------------------
'
'                                   FIELDS
'
'------------------------------------------------------------------------------
'

Private Function FieldProblem( _
    ByRef Value As String, _
    ByRef Column As CsvColumn, _
    ByRef Code As String, _
    ByRef Message As String) _
    As Boolean
'
'==============================================================================
'                                FieldProblem
'------------------------------------------------------------------------------
' PURPOSE
'   Check one filled field against its column, in the reference's order:
'   leading or trailing spaces (all types but text), a formula character
'   at the start (all types but numbers), the length (identifiers and
'   text, in characters), "|" and "#" (identifiers), control characters
'   (line breaks are allowed in text), then the type's own form, allowed
'   values and range.
'
' RETURNS
'   True with Code and Message for the first problem; False when none.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Kind        As String     'The column's type
    Dim Limit       As Long       'Length limit; 0 for none
    Dim First       As String     'First character

'------------------------------------------------------------------------------
' CHECKS COMMON TO THE TYPES
'------------------------------------------------------------------------------
        Code = vbNullString
        Message = vbNullString
        Kind = Column.Kind
        First = Left$(Value, 1)
        If Kind <> "text" And (IsSpace(First) Or IsSpace(Right$(Value, 1))) Then
            Code = CSV_TYPE
            Message = "leading or trailing spaces"
        ElseIf Kind <> "number" And Kind <> "integer" And _
               (First = "=" Or First = "+" Or First = "-" Or First = "@" Or First = vbTab Or First = vbCr) Then
            Code = CSV_FORMULA
            Message = "text starts with a formula character"
        Else
            If Kind = "id" Then Limit = CSV_ID_MAX_LENGTH
            If Kind = "text" Then Limit = CSV_TEXT_MAX_LENGTH
            If Limit > 0 And CharacterCount(Value) > Limit Then
                Code = CSV_LENGTH
                Message = CharacterCount(Value) & " characters, at most " & Limit
            ElseIf Kind = "id" And HasAnyOf(Value, CSV_ID_FORBIDDEN) Then
                Code = CSV_TYPE
                Message = "identifier contains | or #"
            ElseIf HasControl(Value, Kind = "text") Then
                Code = CSV_TYPE
                Message = "control character"
            End If
        End If
        If Len(Code) > 0 Then
            FieldProblem = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' THE TYPE'S OWN CHECKS
'------------------------------------------------------------------------------
        Select Case Kind
            Case "code", "flag"
                If Not IsAllowed(Value, Column.Allowed) Then
                    Code = CSV_ALLOWED
                    Message = "'" & Value & "' is not one of " & Replace(Column.Allowed, "|", ", ")
                End If
            Case "currency"
                If Not IsCurrencyCode(Value) Then
                    Code = CSV_TYPE
                    Message = "not a 3-letter code"
                End If
            Case "integer", "number"
                NumberProblem Value, Column, Code, Message
            Case "date"
                DateProblem Value, Code, Message
        End Select
        FieldProblem = (Len(Code) > 0)

End Function


Private Sub NumberProblem( _
    ByRef Value As String, _
    ByRef Column As CsvColumn, _
    ByRef Code As String, _
    ByRef Message As String)
'
'==============================================================================
'                                NumberProblem
'------------------------------------------------------------------------------
' PURPOSE
'   Check an integer or number: its form (an optional "-", no leading
'   zero, an optional fraction and, for a number, an optional exponent),
'   at most CSV_MAX_DIGITS significant digits, a magnitude a worksheet cell
'   holds, then the column's bounds. The magnitude is judged from the
'   digits and the exponent as text, so that no value overflows. Zeros
'   are skipped by position, never by shortening a string, so that a long
'   run of them costs one pass (#446).
'
' RETURNS
'   Code and Message for the first problem; unchanged when none.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Mantissa        As String     'The digits and point before any exponent, without the sign
    Dim ExponentText    As String     'The exponent's digits and sign; "" for none
    Dim Digits          As String     'Mantissa without the point
    Dim IntegerDigits   As Long       'Digits before the point
    Dim FirstSignificant As Long      'Position of the first non-zero digit in Digits
    Dim LastSignificant As Long       'Position of the last non-zero digit in Digits
    Dim Significant     As String     'Significant digits, without leading or trailing zeros
    Dim Exponent        As Double     'Power of ten of the first significant digit
    Dim Number          As Double     'The value

'------------------------------------------------------------------------------
' FORM AND DIGITS
'------------------------------------------------------------------------------
        If Not IsNumberForm(Value, Column.Kind = "number", Mantissa, ExponentText) Then
            Code = CSV_TYPE
            Message = "'" & Value & "' is not " & IIf(Column.Kind = "integer", "an integer", "a number")
            Exit Sub
        End If
        Digits = Replace(Mantissa, ".", "")
        IntegerDigits = InStr(1, Mantissa & ".", ".", vbBinaryCompare) - 1
        FirstSignificant = FirstNonZero(Digits)
        If FirstSignificant > 0 Then
            LastSignificant = Len(Digits)
            Do While Mid$(Digits, LastSignificant, 1) = "0"
                LastSignificant = LastSignificant - 1
            Loop
            Significant = Mid$(Digits, FirstSignificant, LastSignificant - FirstSignificant + 1)
        End If
        If Len(Significant) > CSV_MAX_DIGITS Then
            Code = CSV_TYPE
            Message = "more than " & CSV_MAX_DIGITS & " significant digits"
            Exit Sub
        End If

'------------------------------------------------------------------------------
' MAGNITUDE
'------------------------------------------------------------------------------
        If Len(Significant) > 0 Then
            Exponent = IntegerDigits - FirstSignificant + ExponentValue(ExponentText)
            If Exponent >= LARGEST_EXPONENT Or Exponent < SMALLEST_EXPONENT Or _
               (Exponent = SMALLEST_EXPONENT And _
                StrComp(Left$(Significant & String$(CSV_MAX_DIGITS, "0"), CSV_MAX_DIGITS), _
                        SMALLEST_DIGITS, vbBinaryCompare) <= 0) Then
                Code = CSV_RANGE
                Message = "outside the magnitudes a worksheet cell can hold"
                Exit Sub
            End If
        End If

'------------------------------------------------------------------------------
' BOUNDS
'------------------------------------------------------------------------------
        Number = Val(Value)
        If Column.HasMin And Number < Column.MinValue Then
            Code = CSV_RANGE
            Message = "below " & Column.MinValue
        ElseIf Column.HasMinExclusive And Number <= Column.MinExclusive Then
            Code = CSV_RANGE
            Message = "not above " & Column.MinExclusive
        ElseIf Column.HasMax And Number > Column.MaxValue Then
            Code = CSV_RANGE
            Message = "above " & Column.MaxValue
        End If

End Sub


Private Function IsNumberForm( _
    ByRef Value As String, _
    ByVal AllowFraction As Boolean, _
    ByRef Mantissa As String, _
    ByRef ExponentText As String) _
    As Boolean
'
'==============================================================================
'                                IsNumberForm
'------------------------------------------------------------------------------
' PURPOSE
'   Whether Value is -?(0|[1-9][0-9]*), followed for a number by an
'   optional .[0-9]+ and an optional [eE][+-]?[0-9]+, with ASCII digits
'   only.
'
' RETURNS
'   True with the mantissa without its sign and the exponent's sign and
'   digits.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Position    As Long    'Next character
    Dim Start       As Long    'Start of the mantissa
    Dim Length      As Long    'Characters of Value

'------------------------------------------------------------------------------
' SIGN AND INTEGER PART
'------------------------------------------------------------------------------
        Length = Len(Value)
        Position = 1
        If Mid$(Value, Position, 1) = "-" Then
            Position = Position + 1
        End If
        Start = Position
        If Mid$(Value, Position, 1) = "0" Then
            Position = Position + 1
        ElseIf IsDigit(Mid$(Value, Position, 1)) Then
            Position = SkipDigits(Value, Position)
        Else
            Exit Function
        End If

'------------------------------------------------------------------------------
' FRACTION AND EXPONENT
'------------------------------------------------------------------------------
        If AllowFraction And Mid$(Value, Position, 1) = "." Then
            If Not IsDigit(Mid$(Value, Position + 1, 1)) Then
                Exit Function
            End If
            Position = SkipDigits(Value, Position + 1)
        End If
        Mantissa = Mid$(Value, Start, Position - Start)
        If AllowFraction And (Mid$(Value, Position, 1) = "e" Or Mid$(Value, Position, 1) = "E") Then
            Position = Position + 1
            Start = Position
            If Mid$(Value, Position, 1) = "+" Or Mid$(Value, Position, 1) = "-" Then
                Position = Position + 1
            End If
            If Not IsDigit(Mid$(Value, Position, 1)) Then
                Exit Function
            End If
            Position = SkipDigits(Value, Position)
            ExponentText = Mid$(Value, Start, Position - Start)
        End If
        IsNumberForm = (Position = Length + 1)

End Function


Private Function SkipDigits( _
    ByRef Value As String, _
    ByVal Position As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The position after the ASCII digits that start at Position.
'==============================================================================
'
        SkipDigits = Position
        Do While IsDigit(Mid$(Value, SkipDigits, 1))
            SkipDigits = SkipDigits + 1
        Loop

End Function


Private Function IsDigit( _
    ByVal Character As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Whether Character is one ASCII digit; "" is not.
'==============================================================================
'
        If Len(Character) = 1 Then
            IsDigit = (AscW(Character) >= 48 And AscW(Character) <= 57)
        End If

End Function


Private Function FirstNonZero( _
    ByRef Digits As String) _
    As Long
'
'==============================================================================
' PURPOSE
'   The position of the first digit other than 0; 0 when there is none.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i       As Long    'Digit counter

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        For i = 1 To Len(Digits)
            If Mid$(Digits, i, 1) <> "0" Then
                FirstNonZero = i
                Exit Function
            End If
        Next i

End Function


Private Function ExponentValue( _
    ByVal ExponentText As String) _
    As Double
'
'==============================================================================
' PURPOSE
'   The exponent of a number, "" being 0; any exponent of more than nine
'   digits counts as +/-1E9, which is beyond every limit, so that no
'   exponent overflows.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Negative    As Boolean    'The exponent is negative
    Dim Digits      As String     'Its digits without leading zeros
    Dim Start       As Long       'Position of the first digit other than 0

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
        Negative = (Left$(ExponentText, 1) = "-")
        Start = 1
        If Left$(ExponentText, 1) = "-" Or Left$(ExponentText, 1) = "+" Then
            Start = 2
        End If
        Do While Mid$(ExponentText, Start, 1) = "0"
            Start = Start + 1
        Loop
        Digits = Mid$(ExponentText, Start)
        If Len(Digits) > 9 Then
            ExponentValue = 1000000000#
        ElseIf Len(Digits) > 0 Then
            ExponentValue = Val(Digits)
        End If
        If Negative Then
            ExponentValue = -ExponentValue
        End If

End Function


Private Sub DateProblem( _
    ByRef Value As String, _
    ByRef Code As String, _
    ByRef Message As String)
'
'==============================================================================
'                                 DateProblem
'------------------------------------------------------------------------------
' PURPOSE
'   Check a yyyy-mm-dd date: ASCII digits in that form, a real day of the
'   Gregorian calendar from year 1, then the contract's range, compared as
'   text, which orders yyyy-mm-dd dates correctly.
'
' RETURNS
'   Code and Message for the problem; unchanged when none.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim YearNumber      As Long       'Year
    Dim MonthNumber     As Long       'Month
    Dim DayNumber       As Long       'Day
    Dim i               As Long       'Character counter
    Dim WellFormed      As Boolean    'Digits and dashes in place

'------------------------------------------------------------------------------
' FORM
'------------------------------------------------------------------------------
        WellFormed = (Len(Value) = 10)
        For i = 1 To Len(Value)
            If i = 5 Or i = 8 Then
                If Mid$(Value, i, 1) <> "-" Then WellFormed = False
            ElseIf Not IsDigit(Mid$(Value, i, 1)) Then
                WellFormed = False
            End If
        Next i
        If Not WellFormed Then
            Code = CSV_TYPE
            Message = "'" & Value & "' is not a yyyy-mm-dd date"
            Exit Sub
        End If

'------------------------------------------------------------------------------
' CALENDAR AND RANGE
'------------------------------------------------------------------------------
        YearNumber = Val(Left$(Value, 4))
        MonthNumber = Val(Mid$(Value, 6, 2))
        DayNumber = Val(Mid$(Value, 9, 2))
        If YearNumber < 1 Or MonthNumber < 1 Or MonthNumber > 12 Or DayNumber < 1 Or _
           DayNumber > DaysInMonth(YearNumber, MonthNumber) Then
            Code = CSV_TYPE
            Message = "'" & Value & "' is not a calendar date"
        ElseIf StrComp(Value, CSV_DATE_MIN, vbBinaryCompare) < 0 Or _
               StrComp(Value, CSV_DATE_MAX, vbBinaryCompare) > 0 Then
            Code = CSV_RANGE
            Message = "outside " & CSV_DATE_MIN & " to " & CSV_DATE_MAX
        End If

End Sub


Private Function DaysInMonth( _
    ByVal YearNumber As Long, _
    ByVal MonthNumber As Long) _
    As Long
'
'==============================================================================
' PURPOSE
'   The days of a month of the proleptic Gregorian calendar.
'==============================================================================
'
        Select Case MonthNumber
            Case 4, 6, 9, 11
                DaysInMonth = 30
            Case 2
                If (YearNumber Mod 4 = 0 And YearNumber Mod 100 <> 0) Or YearNumber Mod 400 = 0 Then
                    DaysInMonth = 29
                Else
                    DaysInMonth = 28
                End If
            Case Else
                DaysInMonth = 31
        End Select

End Function


'
'------------------------------------------------------------------------------
'
'                                 CHARACTERS
'
'------------------------------------------------------------------------------
'

Private Function IsSpace( _
    ByVal Character As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Whether Character is one Python removes with str.strip; "" is not.
'==============================================================================
'
        If Len(Character) = 1 Then
            IsSpace = (InStr(1, SPACE_CODES, "|" & (AscW(Character) And &HFFFF&) & "|", vbBinaryCompare) > 0)
        End If

End Function


Private Function HasControl( _
    ByRef Value As String, _
    ByVal AllowLineBreaks As Boolean) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Whether Value holds a control character (Unicode category Cc: U+0000
'   to U+001F and U+007F to U+009F), a line feed or carriage return
'   excepted when AllowLineBreaks.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CodeUnit    As Long    'A UTF-16 code unit
    Dim i           As Long    'Character counter

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For i = 1 To Len(Value)
            CodeUnit = AscW(Mid$(Value, i, 1)) And &HFFFF&
            If CodeUnit <= &H1F Or (CodeUnit >= &H7F And CodeUnit <= &H9F) Then
                If Not (AllowLineBreaks And (CodeUnit = 10 Or CodeUnit = 13)) Then
                    HasControl = True
                    Exit Function
                End If
            End If
        Next i

End Function


Private Function CharacterCount( _
    ByRef Value As String) _
    As Long
'
'==============================================================================
' PURPOSE
'   The number of characters (code points) of Value: a surrogate pair is
'   one character, as in the reference.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CodeUnit    As Long    'A UTF-16 code unit
    Dim i           As Long    'Code unit counter

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        CharacterCount = Len(Value)
        For i = 1 To Len(Value)
            CodeUnit = AscW(Mid$(Value, i, 1)) And &HFFFF&
            If CodeUnit >= &HDC00& And CodeUnit <= &HDFFF& Then
                CharacterCount = CharacterCount - 1
            End If
        Next i

End Function


Private Function HasAnyOf( _
    ByRef Value As String, _
    ByVal Characters As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Whether Value holds any of Characters.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i       As Long    'Character counter

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For i = 1 To Len(Characters)
            If InStr(1, Value, Mid$(Characters, i, 1), vbBinaryCompare) > 0 Then
                HasAnyOf = True
                Exit Function
            End If
        Next i

End Function


Private Function IsAllowed( _
    ByRef Value As String, _
    ByVal Allowed As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Whether Value is exactly one of the "|"-separated Allowed values.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Values  As Variant    'Allowed values
    Dim n       As Long       'Value counter

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        Values = Split(Allowed, "|")
        For n = 0 To UBound(Values)
            If StrComp(Value, CStr(Values(n)), vbBinaryCompare) = 0 Then
                IsAllowed = True
                Exit Function
            End If
        Next n

End Function


Private Function IsCurrencyCode( _
    ByRef Value As String) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Whether Value is three ASCII capital letters.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CodeUnit    As Long    'A UTF-16 code unit
    Dim i           As Long    'Character counter

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        If Len(Value) <> 3 Then
            Exit Function
        End If
        For i = 1 To 3
            CodeUnit = AscW(Mid$(Value, i, 1)) And &HFFFF&
            If CodeUnit < 65 Or CodeUnit > 90 Then
                Exit Function
            End If
        Next i
        IsCurrencyCode = True

End Function


'
'------------------------------------------------------------------------------
'
'                                  RESULTS
'
'------------------------------------------------------------------------------
'

Private Function ColumnIndex( _
    ByRef Spec As CsvSpec, _
    ByRef Columns() As CsvColumn, _
    ByVal ColumnName As String) _
    As Long
'
'==============================================================================
' PURPOSE
'   The contract column named ColumnName, case-sensitive; 0 when none.
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c       As Long    'Column counter

'------------------------------------------------------------------------------
' FIND
'------------------------------------------------------------------------------
        For c = 1 To Spec.ColumnCount
            If StrComp(Columns(c).ColumnName, ColumnName, vbBinaryCompare) = 0 Then
                ColumnIndex = c
                Exit Function
            End If
        Next c

End Function


Private Sub StartResult( _
    ByRef Result As CsvReadResult, _
    ByRef Spec As CsvSpec, _
    ByRef Values() As String, _
    ByRef Lines() As Long, _
    ByRef Findings() As CsvFinding)
'
'==============================================================================
' PURPOSE
'   Empty Result and its arrays for a read of Spec's kind.
'==============================================================================
'
        Result.Kind = Spec.Kind
        Result.Accepted = False
        Result.RecordCount = 0
        Result.ColumnCount = Spec.ColumnCount
        Result.FindingCount = 0
        Erase Values
        Erase Lines
        ReDim Findings(1 To 16)

End Sub


Private Sub AddFinding( _
    ByRef Result As CsvReadResult, _
    ByRef Findings() As CsvFinding, _
    ByVal Code As String, _
    ByVal LineNumber As Long, _
    ByVal RecordNumber As Long, _
    ByVal ColumnName As String, _
    ByVal Message As String)
'
'==============================================================================
' PURPOSE
'   Append one finding, doubling the array when full.
'==============================================================================
'
        Result.FindingCount = Result.FindingCount + 1
        If Result.FindingCount > UBound(Findings) Then
            ReDim Preserve Findings(1 To 2 * UBound(Findings))
        End If
        With Findings(Result.FindingCount)
            .Code = Code
            .LineNumber = LineNumber
            .RecordNumber = RecordNumber
            .ColumnName = ColumnName
            .Message = Message
        End With

End Sub


Private Sub SortFindings( _
    ByRef Result As CsvReadResult, _
    ByRef Findings() As CsvFinding)
'
'==============================================================================
'                                SortFindings
'------------------------------------------------------------------------------
' PURPOSE
'   Sort the findings by record, line, column and code, as the reference
'   does, and trim the array to them. A merge sort keeps equal findings in
'   the order they were found.
'
' UPDATED
'   2026-10-23
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Order()         As Long       'Finding indices in sorted order
    Dim Spare()         As Long       'Work array of the merge
    Dim Codes()         As String     'Codes in sorted order
    Dim LineNumbers()   As Long       'Lines in sorted order
    Dim RecordNumbers() As Long       'Records in sorted order
    Dim ColumnNames()   As String     'Columns in sorted order
    Dim Messages()      As String     'Messages in sorted order
    Dim RunWidth        As Long       'Length of the runs being merged
    Dim Low             As Long       'Start of the left run
    Dim Middle          As Long       'End of the left run
    Dim High            As Long       'End of the right run
    Dim i               As Long       'Left run position
    Dim j               As Long       'Right run position
    Dim k               As Long       'Merged position
    Dim n               As Long       'Number of findings
    Dim TakeLeft        As Boolean    'The next finding comes from the left run

'------------------------------------------------------------------------------
' MERGE RUNS OF GROWING WIDTH
'------------------------------------------------------------------------------
        n = Result.FindingCount
        If n = 0 Then
            Erase Findings
            Exit Sub
        End If
        ReDim Order(1 To n)
        ReDim Spare(1 To n)
        For k = 1 To n
            Order(k) = k
        Next k
        RunWidth = 1
        Do While RunWidth < n
            Low = 1
            Do While Low <= n
                Middle = Low + RunWidth - 1
                If Middle > n Then Middle = n
                High = Low + 2 * RunWidth - 1
                If High > n Then High = n
                i = Low
                j = Middle + 1
                For k = Low To High
                    'VBA evaluates both sides of And and Or, so each index is
                    'tested before it is used.
                    If i > Middle Then
                        TakeLeft = False
                    ElseIf j > High Then
                        TakeLeft = True
                    Else
                        TakeLeft = Not FindingBefore(Findings(Order(j)), Findings(Order(i)))
                    End If
                    If TakeLeft Then
                        Spare(k) = Order(i)
                        i = i + 1
                    Else
                        Spare(k) = Order(j)
                        j = j + 1
                    End If
                Next k
                Low = High + 1
            Loop
            For k = 1 To n
                Order(k) = Spare(k)
            Next k
            RunWidth = RunWidth * 2
        Loop

'------------------------------------------------------------------------------
' REORDER
'------------------------------------------------------------------------------
    'Field by field, through copies, so that no finding is overwritten
    'before it is read.
        ReDim Codes(1 To n)
        ReDim LineNumbers(1 To n)
        ReDim RecordNumbers(1 To n)
        ReDim ColumnNames(1 To n)
        ReDim Messages(1 To n)
        For k = 1 To n
            Codes(k) = Findings(Order(k)).Code
            LineNumbers(k) = Findings(Order(k)).LineNumber
            RecordNumbers(k) = Findings(Order(k)).RecordNumber
            ColumnNames(k) = Findings(Order(k)).ColumnName
            Messages(k) = Findings(Order(k)).Message
        Next k
        For k = 1 To n
            Findings(k).Code = Codes(k)
            Findings(k).LineNumber = LineNumbers(k)
            Findings(k).RecordNumber = RecordNumbers(k)
            Findings(k).ColumnName = ColumnNames(k)
            Findings(k).Message = Messages(k)
        Next k

End Sub


Private Function FindingBefore( _
    ByRef First As CsvFinding, _
    ByRef Second As CsvFinding) _
    As Boolean
'
'==============================================================================
' PURPOSE
'   Whether First sorts strictly before Second: by record, line, column and
'   code, the texts compared as binary strings.
'==============================================================================
'
        If First.RecordNumber <> Second.RecordNumber Then
            FindingBefore = (First.RecordNumber < Second.RecordNumber)
        ElseIf First.LineNumber <> Second.LineNumber Then
            FindingBefore = (First.LineNumber < Second.LineNumber)
        ElseIf StrComp(First.ColumnName, Second.ColumnName, vbBinaryCompare) <> 0 Then
            FindingBefore = (StrComp(First.ColumnName, Second.ColumnName, vbBinaryCompare) < 0)
        Else
            FindingBefore = (StrComp(First.Code, Second.Code, vbBinaryCompare) < 0)
        End If

End Function
