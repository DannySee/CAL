VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "clsPrograms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'Declare private class constants
Private Const strSheet As String = "Programs"
Private Cosnt strDb As String = "UL_Programs"
Private Const iPKey As Integer = 0
Private Const iCst As Integer = 7
Private Const iSDte As Integer = 9

'Declare private class variables
Private dictCol As New Scripting.Dictionary
Private dictType As New Scripting.Dictionary
Private dictIndex As New Scripting.Dictionary


'*******************************************************************************
'Property to return class sheet name constant (Programs)
'*******************************************************************************
Public Property Get Sht() As String

    'Return class sheet name
    Sht = strSheet
End Property


'*******************************************************************************
'Property to return class database table constant (UL_Programs)
'*******************************************************************************
Public Property Get Db() As String

    'Return class database table name
    Db = strDb
End Property


'*******************************************************************************
'Property to return dictionary of class database field indeces(Key) and
'field names(Item).
'*******************************************************************************
Public Property Get Cols() As Scripting.Dictionary

    'Return class database field dictionary
    Set Cols = dictCol
End Property


'*******************************************************************************
'Property to return dictionary of class database field names(Key) and
'field indeces(Item).
'*******************************************************************************
Public Property Get ColIndex() As Scripting.Dictionary

    'Return class database field dictionary
    Set ColIndex = dictIndex
End Property


'*******************************************************************************
'Property to return dictionary of class database field indeces(Key) and
'field type delimiter(Item). Delimiter is single quote(text) or not(number).
'*******************************************************************************
Public Property Get ColType() As Scripting.Dictionary

    'Return class database field type dictionary
    Set ColType = dictType
End Property


'*******************************************************************************
'Initialize sub. Create dictionaries with column index & name keys
'*******************************************************************************
Private Sub Class_Initialize()

    'Assemble dictionary of field indeces(key) and names(item)
    With dictCol
        .Add 0, "PRIMARY_KEY"
        .Add 1, "CUSTOMER_ID"
        .Add 2, "PROGRAM_ID"
        .Add 3, "DAB"
        .Add 4, "SCRIPT"
        .Add 5, "TIMELINESS"
        .Add 6, "TIER"
        .Add 7, "CUSTOMER"
        .Add 8, "PROGRAM_DESCRIPTION"
        .Add 9, "START_DATE"
        .Add 10, "END_DATE"
        .Add 11, "LEAD_VA"
        .Add 12, "LEAD_CA"
        .Add 13, "VEND_AGMT_TYPE"
        .Add 14, "VENDOR_NUM"
        .Add 15, "BILLBACK_FORMAT"
        .Add 16, "COST_BASIS"
        .Add 17, "CUST_AGMT_TYPE"
        .Add 18, "REBATE_BASIS"
        .Add 19, "PRE_APPROVAL"
        .Add 20, "APPROP_NAME"
        .Add 21, "PRN_GRP"
        .Add 22, "PACKET"
        .Add 23, "PACKET_DL"
        .Add 24, "COMMENTS"
    End With

    'Assemble dictionary of field names(key) and indeces(item)
    With dictIndex
        .Add "PRIMARY_KEY", 0
        .Add "CUSTOMER_ID", 1
        .Add "PROGRAM_ID", 2
        .Add "DAB", 3
        .Add "SCRIPT", 4
        .Add "TIMELINESS", 5
        .Add "TIER", 6
        .Add "CUSTOMER", 7
        .Add "PROGRAM_DESCRIPTION", 8
        .Add "START_DATE", 9
        .Add "END_DATE", 10
        .Add "LEAD_VA", 11
        .Add "LEAD_CA", 12
        .Add "VEND_AGMT_TYPE", 13
        .Add "VENDOR_NUM", 14
        .Add "BILLBACK_FORMAT", 15
        .Add "COST_BASIS", 16
        .Add "CUST_AGMT_TYPE", 17
        .Add "REBATE_BASIS", 18
        .Add "PRE_APPROVAL", 19
        .Add "APPROP_NAME", 20
        .Add "PRN_GRP", 21
        .Add "PACKET", 22
        .Add "PACKET_DL", 23
        .Add "COMMENTS", 24
    End With

    'Assemble dictionary of field indeces(key) and type delimiter(item)
    With dictType
        .Add 0, ""
        .Add 1, ""
        .Add 2, "'"
        .Add 3, "'"
        .Add 4, "'"
        .Add 5, "'"
        .Add 6, "'"
        .Add 7, "'"
        .Add 8, "'"
        .Add 9, "'"
        .Add 10, "'"
        .Add 11, "'"
        .Add 12, "'"
        .Add 13, "'"
        .Add 14, ""
        .Add 15, "'"
        .Add 16, "'"
        .Add 17, "'"
        .Add 18, "'"
        .Add 19, "'"
        .Add 20, "'"
        .Add 21, "'"
        .Add 22, "'"
        .Add 23, "'"
        .Add 24, "'"
    End With
End Sub


'*******************************************************************************
'Return multidimensional array of Excel sheet data.
'*******************************************************************************
Function GetXL() As Variant

    'Declare function variables
    Dim stCon as String

    'Save connection string (connection to CAL workbook)
    stCon "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & ThisWorkbook.FullName & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"

    'Query file (from passthrough sheet) and return results in an open recordset
    rst.Open "SELECT * FROM [" & strSheet & "$] ORDER BY PRIMARY_KEY", stCon

    'Parse recordset into an multidimensional array
    If Not rst.EOF Then var = rst.GetRows()

    'Close recordset and connection & free objects
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing

    'Return multidimensional array of Excel data (from passthrough sheet)
    GetXL = var
End Function


'*******************************************************************************
'Returns updated dictionary of Excel data (Key = Primary_Key, Value = Array
'of fields). Meant to update the passthrough dictionary with static dictionary.
'*******************************************************************************
Function RefreshDct() As Scripting.Dictionary

    'Declare function variables
    Dim dct As New Scripting.Dictionary
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iPKey As Integer
    Dim arr As Variant
    Dim var As Variant

    'Save customer column index (primary key is always first index)
    iPKey = 0

    'Save multidimensional array of program data
    var = GetXL

    'Create an empty array with an index for each program field'
    ReDim arr(UBound(var, 1))

    'Loop through each row of program data
    For iRow = 0 To UBound(var, 2)

        'Loop through each column of program data & add element to array
        For iCol = 0 To UBound(var, 1)
            arr(iCol) = var(iCol, iRow)
        Next

        'Add line to dictionary with program ID as key & row fields as value
        dct(var(iPKey, iRow)) = arr
    Next

    'Return dictionary of program data
    Set RefreshDct = dct
End Function


'*******************************************************************************
'Returns a static dictionary of Program data (Key = Primary Key, Value = array
'of fields). Boolean parameter indicates if the dictionary needs to be updated
'before it is returned.
'*******************************************************************************
Function GetSaveData(blUpdate As Boolean) As Scripting.Dictionary

    'declare static dictionary to hold program data
    Static dct As New Scripting.Dictionary

    'Update dictionary before its returned (if indicated by passthrough boolean)
    If blUpdate Then Set dct = RefreshDct

    'Return static dictionary
    Set GetSaveData = dct
End Function


'*******************************************************************************
'Return multidimensional array of program data that was updated since the Static
'dictionary was initialized. The first index of the return array contains
'program elements to be updated. The second index contains program elements to
'be inserted. Parameters are the static dictionary with historical program data
'and a multidimensional array of current program data.
'*******************************************************************************
Function GetNewRecords() As Scripting.Dictionary

    'Declare function variables
    Dim old As New Scripting.Dictionary
    Dim upd As Variant
    Dim i As Integer
    Dim iSDte As Integer
    Dim iCst As Integer
    Dim iPKey As Integer
    Dim pKey As String
    Dim strInsert As String
    Dim strAssemble As String
    Dim strUpdate As String
    Dim strVal As String
    Dim strInsRows As String

    'Get static dictionary data (historical data for comparison)
    Set old = GetSaveData(False)

    'Get new dictionary data (new data for comparison)
    Set upd = GetXL

    'Loop through rows of current program data
    For iRow = 0 To UBound(upd, 2)

        'If row is new with at least one field filled out (customer name)
        If IsNull(upd(iPKey, iRow)) Then
            If upd(iCst, iRow) <> "" Then

                'Save SQL insert string (including each element of programs tab)
                strInsert = Append(strInsert, "|", GetPrgmInsert(upd, iRow))

                'Save array of insert rows
                strInsRows = Append(strInsRows, "|", iRow + 2)
            End If
        Else


        'Loop through element of doctionary item array
        For i = 0 To UBound(upd(Key))



            'Get current line program ID (used for SQL string)
            pKey = upd(iPKey, iRow)

            'Set temp update SQL string to nothing
            strAssemble = ""

            'Loop through columns of current program data
            For iCol = 0 To UBound(upd, 1)

                'If current data does not match static dictionary
                If (old(pKey)(iCol) <> upd(iCol, iRow)) Or _
                    (IsNull(old(pKey)(iCol)) <> IsNull(upd(iCol, iRow))) Then

                    'Get validated updated value
                    strVal = Validate(upd(iCol, iRow), iCol, iRow, True)

                    'If update is datetime and entry is valid
                    If strVal <> "'DateErr'" Then

                        'Add entry to udate SQL string
                        strAssemble = Append(strAssemble, ",", _
                            oPrgms.Cols(iCol) & " = " & strVal)
                    End If
                End If
            Next

            'If SQL update string contains start date change
            If InStr(strAssemble, "START_DATE") <> 0 Then

                    'Save SQL update string to end date of previous record
                    strUpdate = Append(strUpdate, "|", "END_DATE = '" _
                        & upd(iSDte, iRow) - 1 & "' " _
                        & "WHERE PRIMARY_KEY = " _
                        & upd(iPKey, iRow))

                    'Save SQL insert string
                    strInsert = Append(strInsert, "|", _
                        & GetPrgmInsert(upd, iRow))

                    'Save an array of insert rows
                    strInsRows = Append(strinsRows, "|", iRow + 2)

            'If SQL update string does not contain start date change
            ElseIf strAssemble <> "" Then

                'Save SQL update string
                strUpdate = Append(strUpdate, "|", strAssemble _
                    & " WHERE PRIMARY_KEY = " & upd(iPKey, iRow))
            End If
        End If
    Next

    'If update or insert string is blank (no changes) indicate blank w/ 0 value
    If strUpdate = "" Then strUpdate = "0"
    If strInsert = "" Then strInsert = "0"

    'Return multidimensional array w/ update(0) and insert(1) SQL string arrays
    GetNewRecords = Array(Split(strUpdate, "|"), Split(strInsert, "|"), _
        Split(strInsRows, "|"))
End Function
