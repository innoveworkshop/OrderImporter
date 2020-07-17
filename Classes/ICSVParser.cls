VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICSVParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' CLASS PURPOSE: Parse CSV files and allow user to respond to each record parsed
' USAGE:
'   1) Declare instance of this class at top of your form/class using: WithEvents
'       i.e., Private WithEvents CSVparser As ICSVParser
'   2) Initiate: CSVParser.InitializeParser and pass user-defined options
'   3) Read entire file into memory or read single lines from the CSV in a loop
'       pass that read information to CSVParser.ParseRecord
'   4) Respond to the only event this class passes. Parsed records are passed to that event.
'       That event allows you to abort processing, it also where you would use the parsed
'       field information, i.e., listview, ADO recordset, etc, etc.
'   5) Clean up: CSVParser.TerminateParser

Public Enum ProcessStateEnum
    csvRecordParsed = 0
    csvFieldNamesStatic = 1
    csvFieldNamesGeneric = 2
End Enum

Public Event ProcessRecord(ByVal State As ProcessStateEnum, Record As String, _
                            ByVal FieldDifferential As Long, ByVal HeaderCount As Long, _
                            ByVal RecordNumber As Long)
' the above event is triggered each time a complete CSV record has been processed
' The HeaderCount parameter will always return the number of fields expected in the CSV
' The RecordNumber will be zero when header row is passed & increments at that point for each call
' State = csvRecordParsed
'   Record parameter = vbNullChar delimeted CSV record
'   FieldDifferential parameter is difference from number of fields in record vs header
'   zero indicates match, < 0 indicates lesss records, > 0 indicates more records
' State = csvFieldNamesStatic
'   Record parameter = vbNullChar delimeted Field names. HeaderCount parameter is number of fields
'   The field names are retrieved from the CSV file. FieldDifferential will be zero
' State = csvFieldNamesGeneric
'   Record parameter = vbNullChar delimeted Field names. HeaderCount parameter is number of fields
'   The field names are like: Field 1, Field 2, etc. You should change these, at this time,
'       to any preferred format you wish. FieldDifferential will be zero
' Notes: If during any event, the FieldDifferential is non-zero, it is likely the CSV format
'   is corrupted or possibly that trailing empty fields are not delimited (technically corrupted)
'   Strongly recommend aborting processing... If not, you have a couple of choices:
'   -- if FieldDifferential > 0 then increase your header row for the extra fields
'   -- if FieldDifferential < 0 then know you will not 'fill up' your record
'   -- To abort processing, return the Record parameter as a null/empty string... Example
'       **********************************************************************************
'       On Error GoTo ExitRoutine
'       ' open target CSV file as file number x
'       myCSVParser.InitializeParser True
'       Do Until EOF(x) = True
'           Line Input #x, strRecord
'           If myCSVparser.ParseRecord(strRecord) = False Then
'               ' aborted from myCSVparser_ProcessRecord, inform user of problem
'               Exit Do
'           End If
'       Loop
'       Call myCSVParser.TerminateParser    ' handle any malformatted final record
'
'       ExitRoutine:
'       If Err Then
'           ' notify user of possible corrupt or invalid CSV format
'       End If
'       Close #x
'       **********************************************************************************
'       Private Sub myCSVparser_ProcessRecord(ByVal State As ProcessStateEnum, Record As String, ByVal FieldDifferential As Long, ByVal HeaderCount As Long, ByVal RecordNumber as Long)
'           Select Case State
'           Case csvRecord
'               If FieldDifferential Then
'                   Record = vbNullString   ' <<< this forces ParseRecord() to return False
'               Else
'                   ' vbNullChar delimited CSV record. Split/Process as needed
'               End If
'           Case csvFieldNameStatic
'               ' vbNullChar delimited CSV header row. Split/Process as needed
'           Case csvFieldNameGeneric
'               ' Split/Process as needed. May want to replace with custom field names
'           End Select
'       End Sub
'       **********************************************************************************

Private m_DelimFld As Long          ' user-provided field delimeter. Default is: ,
Private m_DelimQuote As Long        ' user-provided quote delimeter. Default is: "
Private m_DelimQuoteStr As String   ' string version of m_DelimQuote
Private m_NullFieldValue As String  ' user-provided default value for null fields: Default: null string
Private m_Escape As Long            ' optional Escape character
Private m_EscapeStr As String       ' string version of m_Escape
Private m_Fields As Long            ' number of fields processed from 1st record
Private m_Tokens As Long            ' number of fields processed from each record
Private m_Record As String          ' a reusable string holding the current record
Private m_RecordPos As Long         ' write-position within the m_Record string
Private m_RecordCount As Long       ' number of records processed
Private m_State As Long             ' flags used during processing records & fields
    '   0x0001   start of token is with a quote
    '   0x0002   currently processing stuff within quotes
    '   0x0004   last token ended on delimiter (at least 1 more token to be processed)
    '   0x0008   start of token determined
    '   0x0010   cached portion of token includes leading quote to be stripped off
    '   0x0020   cached portion of token includes trailing quote to be stripped off
    '   0x0040   escape character (next character is literal)
    '   0x0080   partial token being written
    '   values > 0xFFF are relative to entire file else relative to current record
    '   0x1000   has header row
    '   0x2000   trim leading/trailing spaces

Public Sub InitializeParser(ByVal HasHeaderRow As Boolean, _
                            Optional FieldDelimeter As String = ",", _
                            Optional QuoteDelimeter As String = """", _
                            Optional EscapeCharacter As String = vbNullString, _
                            Optional NullFieldValue As String = vbNullString, _
                            Optional TrimLeadingTrailingSpaces As Boolean = False)
    
    ' must be called before any CSV is initially processed
    ' Parameters:
    '   HasHeaderRow: If passed as True, the 1st processed record is forwarded as the header row
    '       Otherwise, the header row is created as static field names like Field 1, Field 2, etc
    '   FieldDelimiter: Cannot be null else ParseRecord method will abort
    '       Can use any character you wish, comma is common
    '   QuoteDelimiter: Must be provided if the CSV is formatted with quote delimiters
    '       If not provided, quotes are treated no differently than any other character
    '       More on this a tad later
    '   EscapeCharacter: Optional. Not very common at all.
    '       When provided, the Escape character is not included in parsed data and the
    '       character immediately following the EscapeCharacter is treated as any other character.
    '       To escape an escape character, double them up.
    '   NullFieldValue: Optional. What to pass as the CSV field value if it is blank.
    '   TrimLeadingTrailingSpaces: If true, then VB's Trim() function will be applied
    
    ' Quote Delimited CSVs
    '   Typically, if a CSV field contains a field delimiter, then the entire field will use
    '       a quote as both a prefix/suffix for the field, i.e., "Yes, you can"
    '   It is not uncommon for every 'text datatype' field of a CSV to be delimited by quotes
    '   It is not uncommon to quote-delimit fields that contain leading/trailing spaces
    '   It is standard to quote-delimit fields that contain carriage returns or non-printable characters
    '   Ok, how are quotes within a quote-delimited field handled?
    '       Properly: each quote becomes a double quote. This does not apply to the prefix/suffix
    '       i.e., Print #n, Chr$(34); Replace(myData, Chr$(34), Chr$(34) & Chr$(34); Chr$(34);
    ' Escape Characters
    '   Very similar to quoted fields. Escapes can be used within quoted fields or not
    '   Here are some examples, let's say our escape character is: \
    '       As in CSV           After Processing    Notes
    '       Yes\, you can       Yes, you can
    '       "Yes\, you can"     Yes, you can        QuoteDelimiter = "
    '                           "Yes, you can"      QuoteDelimiter not provided
    '       C:\Temp             C:Temp
    '       C:\\Temp            C:\Temp
    
    ' Note that typically, either Quotes or Escape characters are used, not common for both to be used
    
    m_State = 0&: Call Me.TerminateParser
    If Not FieldDelimeter = vbNullString Then m_DelimFld = AscW(Left$(FieldDelimeter, 1))
    If QuoteDelimeter = vbNullString Then
        QuoteDelimeter = -1&        ' quotes are handled as just any other character
    Else
        m_DelimQuote = AscW(Left$(QuoteDelimeter, 1))
        m_DelimQuoteStr = ChrW$(m_DelimQuote)
    End If
    If EscapeCharacter = vbNullString Then
        m_Escape = -1&              ' no optional Escape character
    Else
        m_Escape = AscW(Left$(EscapeCharacter, 1))
        m_EscapeStr = ChrW$(m_Escape)
    End If
    m_NullFieldValue = NullFieldValue
    If HasHeaderRow Then m_State = &H1000
    If TrimLeadingTrailingSpaces Then m_State = m_State Or &H2000
End Sub

Public Function TerminateParser() As Boolean
    ' Should be called after finishing processing the CSV file.
    ' If the final record did not complete, it is still cached in this class.
    ' By calling this method, the class will forward whatever it has cached and return False
    ' True return value means no left-over data found and/or entire file processed ok
    
    If (m_State And &HFF) Then  ' after a full record processed, these bytes are always zero
        TerminateParser = pvNotifyClient()
    Else
        TerminateParser = True
    End If
    ' Clear all class-level variables
    m_DelimFld = 0&: m_DelimQuote = 0&
    m_DelimQuoteStr = vbNullString
    m_NullFieldValue = vbNullString
    m_Record = vbNullString
    m_RecordPos = 0&: m_State = 0&
    m_Fields = 0&: m_Tokens = 0&
    m_Escape = 0&: m_EscapeStr = vbNullString
    m_RecordCount = 0&
End Function

Public Function ParseRecord(CSVrecord As String) As Boolean

    ' The passed parameter must be one of these two cases
    '   1) Entire CSV file. Must have rows delimited by both or either: vbCr, vbLf
    '   2) One or more complete lines of the CSV file, as within a loop, reading line by line
    '       The line, in this case, is assumed terminated by a carriage return or line feed
    '       and therefore, it is not necessary to include carriage returns for single lines
    ' This parser fails in these three specific cases
    '   1) InitializeParser not called
    '   2) Passed record contains vbNullChar (ASCII byte 0)
    '   3) User makes the ProcessRecord event return null Record parameter; aborted by user

    If m_DelimFld = 0& Then Exit Function ' InitializeParser not called

    Dim c As Long, lChar As Long, lStart As Long
    
    If CSVrecord = vbNullString Then
        ' empty string = carriage return only
        ' if processing quoted text, add this to the record
        If (m_State And 8) Then pvAddField vbCrLf, 1&, 3&, False
        ParseRecord = True
        Exit Function
    End If
    
    For c = 1 To Len(CSVrecord)
        lChar = AscW(Mid$(CSVrecord, c, 1))
        If lChar = m_Escape Then
            If (m_State And 8) = 0& Then    ' start of field?
                lStart = c: m_State = (m_State And &HFFFFFF00) Or 8&
            End If
            c = c + 1&                     ' move pointer along & adjust if necessary
            If c > Len(CSVrecord) Then c = Len(CSVrecord)
        
        '/// Check #2: Field delimeter?
        ElseIf lChar = m_DelimFld Then
            If (m_State And 2) = 0 Then     ' not inside quoted text
                If (m_State And 8) = 0& Then    ' no field data
                    pvAddField vbNullString, 0&, 0&, ((m_State And &H80) = 0&)
                Else
                    pvAddField CSVrecord, lStart, c, ((m_State And &H80) = 0&)
                End If                          ' ensure m_State not null
                m_State = (m_State And &HFFFFFF00) Or 4&
            End If
        
        '/// Check #3: Processing quoted field?
        ElseIf lChar = m_DelimQuote Then
            If (m_State And 8) = 0& Then
                lStart = c: m_State = (m_State And &HFFFFFF00) Or &H1B  ' &H1 or &H2 or &H8 or &H10
            Else
                m_State = m_State Xor 2         ' toggle quote flag
                If (m_State And 1) Then m_State = m_State Xor &H20 ' toggle whether trailing quote is applicable
            End If
            
        '/// Check #4: Record delimeter?
        ElseIf lChar = 10 Or lChar = 13 Then
            If (m_State And &HE) > 2& Then      ' else inside quoted text or not processing field
                If (m_State And 8) = 0& Then
                    pvAddField vbNullString, 0&, 0&, True
                Else
                    pvAddField CSVrecord, lStart, c, True
                End If
                If pvNotifyClient() = False Then Exit Function
            End If
            
        '/// Check #5: Start of record found?
        ElseIf (m_State And 8) = 0& Then
            Select Case lChar
            Case 32 ' white space. Add other character codes as needed
                If (m_State And &H2000) = 0& Then
                    lStart = c: m_State = (m_State And &HFFFFFF00) Or 8&
                End If
            Case Is < 0, Is > 31
                lStart = c: m_State = (m_State And &HFFFFFF00) Or 8&
            Case 0: Exit Function   ' abort
            Case Else               ' Is < 32; control characters?
                ' ignored as start of a field
            End Select
            
        ElseIf lChar = 0& Then
            Exit Function           ' abort
        End If
        
    Next
    
    ' finish off the passed field data
    If (m_State And 2) Then                 ' within quoted text, add this parital to the field data
        pvAddField CSVrecord, lStart, c, ((m_State And &H80) = 0)
        m_State = (m_State And Not &H10) Or &H80 ' flag for partial record & remove flag for leading quote
        pvAddField vbCrLf, 1, 3, False      ' append carriage return
        ParseRecord = True
        
    ElseIf (m_State And &HFF&) Then
        If (m_State And 8) = 0 Then         ' no field data
            pvAddField vbNullString, 0&, 0&, ((m_State And &H80) = 0&)
        Else                                ' else add field data
            pvAddField CSVrecord, lStart, c, ((m_State And &H80) = 0&)
        End If
        ParseRecord = pvNotifyClient()      ' done with the record
    
    Else
        ParseRecord = True
    End If

End Function

Private Function pvNotifyClient() As Boolean
    
    ' inform client a record has been processed
    
    Dim sRecord As String
    If m_Fields = 0& Then                    ' 1st record, no header processed yet
        m_Fields = m_Tokens
        If (m_State And &H1000&) = 0 Then    ' no header? provide default field names
            For m_Tokens = 1& To m_Tokens - 1&
                sRecord = sRecord & "Field " & CStr(m_Tokens) & vbNullChar
            Next                            ' forward the event
            sRecord = sRecord & "Field " & CStr(m_Tokens)
            RaiseEvent ProcessRecord(csvFieldNamesGeneric, sRecord, 0&, m_Fields, m_RecordCount)
            If Not sRecord = vbNullString Then
                pvNotifyClient = pvNotifyClient()
                Exit Function
            End If
        Else                                ' forward the header row
            sRecord = Left$(m_Record, m_RecordPos)
            RaiseEvent ProcessRecord(csvFieldNamesStatic, sRecord, 0&, m_Fields, m_RecordCount)
        End If
    Else
        If m_Tokens = 0& Then               ' prevent passing null string
            sRecord = vbNullChar
        Else                                ' extract current record
            sRecord = Left$(m_Record, m_RecordPos)
        End If                              ' pass to user & check for abort flag
        RaiseEvent ProcessRecord(csvRecordParsed, sRecord, m_Tokens - m_Fields, m_Fields, m_RecordCount)
    End If
    m_Tokens = 0&: m_RecordPos = 0&         ' reset record-related flags & indexes
    m_State = m_State And &HFFFFFF00
    m_RecordCount = m_RecordCount + 1&
    pvNotifyClient = Not (sRecord = vbNullString)
    
    
End Function

Private Sub pvAddField(FieldValue As String, StartOffset As Long, EndOffset As Long, AddDelimeter As Boolean)

    ' add the passed field info to the current record

    Dim lSize As Long, sRecord As String
    
    If StartOffset > EndOffset Then EndOffset = 0&  ' sanity check; should never occur
    
    If EndOffset Then                       ' else null field data
        If StartOffset = 0& Then StartOffset = 1&
        If (m_State And &H2000) Then    ' trim leading/trailing spaces?
            ' but must be ignored, entirely or in part, if saving a partial record within quoted text
            If (m_State And 2) Then         ' within quoted text
                If AddDelimeter Then        ' saving 1st portion of partial record, trim left edge
                    sRecord = LTrim$(Mid$(FieldValue, StartOffset, EndOffset - StartOffset))
                Else                        ' no trimming at all
                    sRecord = Mid$(FieldValue, StartOffset, EndOffset - StartOffset)
                End If
            ElseIf AddDelimeter = False Then ' saving final portion of record, trim right edge
                sRecord = RTrim$(Mid$(FieldValue, StartOffset, EndOffset - StartOffset))
            Else
                sRecord = Trim$(Mid$(FieldValue, StartOffset, EndOffset - StartOffset))
            End If
        Else                                ' no trim option
            sRecord = Mid$(FieldValue, StartOffset, EndOffset - StartOffset)
        End If
        If Not m_DelimQuote = -1& Then      ' replace double quotes with quote
            If InStr(sRecord, m_DelimQuoteStr & m_DelimQuoteStr) Then
                sRecord = Replace(sRecord, m_DelimQuoteStr & m_DelimQuoteStr, m_DelimQuoteStr)
            End If
        End If
        If Not m_Escape = -1& Then          ' remove escape characters after replacing double escapes
            If InStr(sRecord, m_EscapeStr) Then
                If InStr(sRecord, m_EscapeStr & m_EscapeStr) Then
                    sRecord = Replace(sRecord, m_EscapeStr & m_EscapeStr, Chr$(8))
                    sRecord = Replace(sRecord, m_EscapeStr, vbNullString)
                    sRecord = Replace(sRecord, Chr$(8), m_EscapeStr)
                Else
                    sRecord = Replace(sRecord, m_EscapeStr, vbNullString)
                End If
            End If
        End If
        StartOffset = 1&: lSize = Len(sRecord)
        If (m_State And &H30) Then          ' handle leading/trailing quote
            If (m_State And &H10) Then StartOffset = StartOffset + 1: lSize = lSize - 1&
            If (m_State And &H20) Then lSize = lSize - 1&
        End If
    
    ElseIf m_Fields = 0& And (m_State And &H1000&) Then ' deal with null header field
        ' if header row exists & an empty field is retrieved, create a unique header field
        sRecord = "Unknown " & CStr(m_Tokens + 1&)
        lSize = Len(sRecord): StartOffset = 1&
    
    Else
        lSize = Len(m_NullFieldValue)
    End If
    
    If Len(m_Record) < lSize + m_RecordPos + 1& Then ' resize general use string larger than needed
        m_Record = m_Record & Space$(500& + lSize)
    End If
    If m_RecordPos > 0& Then                        ' not the 1st field for this record
        If AddDelimeter Then                        ' append delimeter?
            m_RecordPos = m_RecordPos + 1&
            Mid$(m_Record, m_RecordPos, 1&) = vbNullChar
        End If
    End If
    If EndOffset Then                               ' non-null field
        Mid$(m_Record, m_RecordPos + 1&, lSize) = Mid$(sRecord, StartOffset, lSize)
    ElseIf lSize Then
        Mid$(m_Record, m_RecordPos + 1&, lSize) = m_NullFieldValue
    End If
    m_RecordPos = m_RecordPos + lSize               ' update current record size & token count
    If AddDelimeter Then m_Tokens = m_Tokens + 1&
End Sub
