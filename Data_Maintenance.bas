Attribute VB_Name = "Data_Maintenance"


'*******************************************************************************
'Returns a static dicitonary of Program data (Key = Primary Key, Value = array
'of fields). Boolean operator indicates if the dictionary needs to be
'updated before it is returned.
'*******************************************************************************
Function dctPrograms(blUpdate As Boolean) As Scripting.Dictionary

    'declare static dictionary to hold program data
    Static dctPrgms As New Scripting.Dictionary

    'Update dictionary before its returned (if indicated by passthrough boolean)
    If blUpdate Then Set dctPrgms = UpdateDct(dctPrgms, "Programs")

    'Return static dictionary
    Set dctPrograms = dctPrgms
End Function


'*******************************************************************************
'Returns a static dicitonary of Customer Profile data (Key = Primary Key, Value
'= array of fields). Boolean operator indicates if the dictionary needs to
'be updated before it is returned.
'*******************************************************************************
Function dctCstProfile(blUpdate As Boolean) As Scripting.Dictionary

    'declare static dictionary to hold program data
    Static dctCst As New Scripting.Dictionary

    'Update dictionary before its returned (if indicated by passthrough boolean)
    If blUpdate Then Set dctCst = UpdateDct(dctCst, "Customer Profile")

    'Return static dictionary
    Set dctPrograms = dctCst
End Function


'*******************************************************************************
'Returns updated dictionary of Excel data (Key = Primary_Key, Value = Array
'of fields). Meant to update the passthrough static dictionary from
'passthrough sheet.
'*******************************************************************************
Function UpdateDct(dct As Scripting.Dictionary, strSht As String) As _
    Scripting.Dictionary

    'Declare function variables
    Dim iRow, iCol As Integer
    Dim iPKey As Integer
    Dim arr As Variant
    Dim var As Variant

    'Save customer column index (primary key is always first index)
    iPKey = 0

    'Clear values items
    dct.RemoveAll

    'Save multidimensional array of program data
    var = GetXL(strSht & "$")

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
    Set UpdateDct = dct
End Function


'*******************************************************************************
'Return multidimensional array of Excel file. Pass through string indicates
'tab's data should be returned.
'*******************************************************************************
Function GetXL(strSheet As String) As Variant

    'Declare function variables
    Dim stCon as String

    'Save connection string (connection to CAL workbook)
    stCon "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & ThisWorkbook.FullName & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"

    'Query file (from passthrough sheet) and return results in an open recordset
    rst.Open "SELECT * FROM [" & strSheet & "]", stCon

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
'Return multidimensional array of program data that was updated since the Static
'dictionary was initialized. The first index of the return array contains
'program elements to be updated. The second index contains program elements to
'be inserted. Passthrough variables are the static dictionary with historical
'program data and a multidimensional array of current program data.
'*******************************************************************************
Function ComparePrgms(old As Scripting.Dictionary, upd As Variant) As Variant

    'Declare function variables
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iSDateCol As Integer
    Dim iCustCol As Integer
    Dim iPKeyCol As Integer
    Dim pKey As String
    Dim strInsert As String
    Dim strAssemble As String
    Dim strUpdate As String
    Dim strVal As String
    Dim strInsRows As String

    'Get column index for pertinant fields
    iCustCol = oPrgms.ColIndex("CUSTOMER")
    iSDateCol = oPrgms.ColIndex("START_DATE")
    iPKeyCol = oPrgms.ColIndex("PRIMARY_KEY")

    'Loop through rows of current program data
    For iRow = 0 To UBound(upd, 2)

        'If row is new with at least one field filled out (customer name)
        If IsNull(upd(iPKeyCol, iRow)) Then
            If upd(iCustCol, iRow) <> "" Then

                'Save SQL insert string (including each element of programs tab)
                strInsert = Append(strInsert, "|", GetPrgmInsert(upd, iRow))

                'Save array of insert rows
                strInsRows = Append(strInsRows, "|", iRow + 2)
            End If
        Else

            'Get current line program ID (used for SQL string)
            pKey = upd(iPKeyCol, iRow)

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
                        & upd(iSDateCol, iRow) - 1 & "' " _
                        & "WHERE PRIMARY_KEY = " _
                        & upd(iPKeyCol, iRow))

                    'Save SQL insert string
                    strInsert = Append(strInsert, "|", _
                        & GetPrgmInsert(upd, iRow))

                    'Save an array of insert rows
                    strInsRows = Append(strinsRows, "|", iRow + 2)

            'If SQL update string does not contain start date change
            ElseIf strAssemble <> "" Then

                'Save SQL update string
                strUpdate = Append(strUpdate, "|", strAssemble _
                    & " WHERE PRIMARY_KEY = " & upd(iPKeyCol, iRow))
            End If
        End If
    Next

    'If update or insert string is blank (no changes) indicate blank w/ 0 value
    If strUpdate = "" Then strUpdate = "0"
    If strInsert = "" Then strInsert = "0"

    'Return multidimensional array w/ update(0) and insert(1) SQL string arrays
    Compare = Array(Split(strUpdate, "|"), Split(strInsert, "|"), _
        Split(strInsRows, "|"))
End Function


'*******************************************************************************
'Return multidimensional array of program data that was updated since the Static
'dictionary was initialized. The first index of the return array contains
'program elements to be updated. The second index contains program elements to
'be inserted. Passthrough variables are the static dictionary with historical
'program data and a multidimensional array of current program data.
'*******************************************************************************
Function CompareCst(old As Scripting.Dictionary, upd As Variant) As Variant

    'Declare function variables
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iCust As Integer
    Dim iPKey As Integer
    Dim pKey As String
    Dim strInsert As String
    Dim strAssemble As String
    Dim strUpdate As String
    Dim strVal As String
    Dim strInsRows As String

    'Get column index for pertinant fields
    iPKey = oPrgms.ColIndex("PRIMARY_KEY")
    iCust = oPrgms.ColIndex("CUSTOMER")

    'Loop through rows of current program data
    For iRow = 0 To UBound(upd, 2)

        'If row is new with at least one field filled out (customer name)
        If IsNull(upd(iPKey, iRow)) Then
            If upd(iCust, iRow) <> "" Then

                'Save SQL insert string (including each element of programs tab)
                strInsert = Append(strInsert, "|", GetCstInsert(upd, iRow))

                'Save array of insert rows
                strInsRows = Append(strInsRows, "|", iRow + 2)
            End If
        Else

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
                    strVal = Validate(upd(iCol, iRow), iCol, iRow, False)

                    'If update is datetime and entry is valid
                    If strVal <> "'DateErr'" Then

                        'Add entry to udate SQL string
                        strAssemble = Append(strAssemble, ",", _
                            oCst.Cols(iCol) & " = " & strVal)
                    End If
                End If
            Next

            'Save SQL update string
            strUpdate = Append(strUpdate, "|", strAssemble _
                & " WHERE PRIMARY_KEY = " & upd(iPKey, iRow))
        End If
    Next

    'If update or insert string is blank (no changes) indicate blank w/ 0 value
    If strUpdate = "" Then strUpdate = "0"
    If strInsert = "" Then strInsert = "0"

    'Return multidimensional array w/ update(0) and insert(1) SQL string arrays
    Compare = Array(Split(strUpdate, "|"), Split(strInsert, "|"), _
        Split(strInsRows, "|"))
End Function


'*******************************************************************************
'Return value which has been validated for its SQL field datatype. This Function
'is meant to assist in assembling SQL update/insert strings to the UL_Programs
'database. Passthrough variables are the string value to be validated and its
'origin row/column. blDataType passthrough indicates if the data type needs to
'be validated. Column/row index is from multidimensional array, not excel
'coordinates.
'*******************************************************************************
Function Validate(val As Variant, iCol As Integer, iRow As Integer, _
    blDatType As Boolean) As String

    'Declare function variables
    Dim sep As String
    Dim iSDateCol As integer
    Dim iEDateCol As integer
    Dim iVendCol As integer

    'If data type needs to be validated
    If blDataType = True Then

        'Get string with datatype delimiters (quotes for text, nothing for number)
        sep = oPrgms.ColType(iCol)

        'Get column index for pertinant fields
        iSDateCol = oPrgms.ColIndex("START_DATE")
        iEDateCol = oPrgms.ColIndex("END_DATE")
        iVendCol = oPrgms.ColIndex("VENDOR_NUM")

        'If passthrough data type is datetime and value is invalid date
        If (iCol = iSDateCol Or iCol = iEDateCol) And Not IsDate(val) Then

            'Alert user of invalid date value
            MsgBox "INVALID DATE: Please correct the " & oPrgms.Cols(iCol) _
                & " field - " & val & " is not a valid date entry"

            'Highlight (select) row with invalid date value to alert user
            Rows(iRow + 2).EntireRow.Select

            'Return invalid date error message
            val = "DateErr"

        'If passthrough datatype is number and value is empty
        ElseIf iCol = iVendCol And IsNull(val) Then

            'Return 0 value (string to maintain function type)
            val = "0"
        End if
    Else

        'Set string delimiter to SQL syntax single quotes
        sep = "'"
    End If

    'Return validated value wrapped in datatype appropriate delimiter
    Validate = sep & Replace(val, "'", "") & sep
End Function


'*******************************************************************************
'Returns a concatenated string and separator.
'*******************************************************************************
Function Append(val As String, sep As String, val2 As String) As String

    'If value is blank
    If val = "" Then

        'Return blank variable
        Append = val2
    Else

        'Return concatenated value and separator
        Append = val & sep & " " & val2
    End If
End Function


'*******************************************************************************
'Returns an SQL insert statement. Function is Specific to programs tab. Gutter
'Data is Primary Key, Cusotmer ID and Program ID. Guts data is all other program
'fields. Passthrough variables are multidimensional array of programs tab and
'focus row index.
'*******************************************************************************
Function GetPrgmInsert(var As Variant, iRow As integer) As String

    'Declare function variables
    Dim iPKey As Integer
    Dim iCID As integer
    Dim iPID As integer
    Dim iDAB As integer
    Dim strGutter As String
    Dim strGuts As String

    'Get column index for pertinant fields
    iPKey = oPrgms.ColIndex("PRIMARY_KEY")
    iCID = oPrgms.ColIndex("CUSTOMER_ID")
    iPID = oPrgms.ColIndex("PROGRAM_ID")
    iDAB = oPrgms.ColIndex("DAB")
    iCust = oPrgms.ColIndex("CUSTOMER")

    'If array does not contain customer ID data
    If IsNull(var(iPKey, iRow)) Then

        'Establish conection to SQL server
        cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1" _
            & ";DATABASE=Pricing_Agreements;"

        'Query customer and program IDs from customer name
        rst.Open "SELECT TOP 1 CUSTOMER_ID AS CID, " _
            & "MAX(CAST(RIGHT(PROGRAM_ID, " _
            & "CHARINDEX('-', REVERSE(PROGRAM_ID)) - 1) AS INT)) + 1 AS PID " _
            & "FROM UL_Programs WHERE CUSTOMER = '" _
            & var(iCust, iRow) & "' " _
            & "GROUP BY CUSTOMER_ID", cnn

        'Assemble first 3 fields of SQL insert string
        strGutter = rst.Fields("CID").value & ",'" & rst.Fields("CID").value _
            & "-" & rst.Fields("PID").value & "'"

        'Close and free objects
        rst.Close
        cnn.Close
        Set rst = Nothing
        Set cnn = Nothing
    Else

        'Assemble first 3 fields of SQL insert string
        strGutter = var(iCID, iRow) & ",'" _
            & var(iPID, iRow) & "'"
    End If

    'Get concatenated string of all fields in SQL sytnax
    strGuts = AppendRow(var, iRow, iDAB, True)

    'Return Programs gutter
    GetPrgmInsert = Append(strGutter, ",", strGuts)
End Function


'*******************************************************************************
'Returns an SQL insert statement. Function is Specific to customer profile tab.
'Gutter data is Primary Key, Cusotmer ID and Program ID. Guts data is all other
'program fields. Passthrough variables are multidimensional array of programs
'tab and focus row index.
'*******************************************************************************
Function GetCstInsert(var As Variant, iRow As integer) As String

    'Declare function variables
    Dim iPKey As Integer
    Dim iCID As integer
    Dim iCust As integer
    Dim strGutter As String
    Dim strGuts As String

    'Get column index for pertinant fields
    iPKey = oCst.ColIndex("PRIMARY_KEY")
    iCID = oCst.ColIndex("CUSTOMER_ID")
    iCust = oCst.ColIndex("CUSTOMER")

    'If array does not contain customer ID data
    If IsNull(var(iPKey, iRow)) Then

        'Establish conection to SQL server
        cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1" _
            & ";DATABASE=Pricing_Agreements;"

        'Query customer and program IDs from customer name
        rst.Open "SELECT TOP 1 CUSTOMER_ID AS CID " _
            & "FROM UL_Customer_Profile WHERE CUSTOMER = '" _
            & var(iCust, iRow) & "' ", cnn

        'Assemble first 3 fields of SQL insert string
        strGutter = rst.Fields("CID").value

        'Close and free objects
        rst.Close
        cnn.Close
        Set rst = Nothing
        Set cnn = Nothing
    Else

        'Assemble first 3 fields of SQL insert string
        strGutter = var(iCID, iRow)
    End If

    'Get concatenated string of all fields in SQL sytnax
    strGuts = AppendRow(var, iRow, iCust, False)

    'Return Programs gutter
    GetPrgmInsert = Append(strGutter, ",", strGuts)
End Function


'*******************************************************************************
'Returns a concatenated string in SQL syntax. Passthrough variables are a
'multidimensional array with Excel data, the row index to be parsed and the
'starting field index. Boolean value is to indicate if validation method
'requires data type check.
'*******************************************************************************
Function AppendRow(var As Variant, iRow As Integer, _
    iStart As Integer, blType) As String

    'Declare function variables
    Dim iCol As Integer
    Dim strVal As String
    Dim strRow As String

    'Loop through each column of passthrough array
    For iCol = iStart To UBound(var, 1)

        'Assemble SQL insert string
        strVal = Validate(var(iCol, iRow), iCol, iRow, blType)

        'Assemble array of insert rows
        strRow = Append(strRow, ",", strVal)
    Next

    'Return SQL insert string
    AppendRow = strRow
End Function


'*******************************************************************************
'Executes SQL update statement to CAL database. Passthrough variable is an array
'of update statements, one element per statement.
'*******************************************************************************
Sub UploadChanges(upd As Variant)

    'Declare function variables
    Dim i As Integer

    'Loop through each update statement in passthrough array
    For i = 0 To UBound(upd)

        'Execute update statement
        cnn.Execute "UPDATE UL_Programs SET " & upd(i)
    Next
End Sub


'*******************************************************************************
'Execute SQL insert statement to CAL database. passthrough variables are an
'array of insert statements (one element per statement) and an array with the
'corresponding excel row numbers.
'*******************************************************************************
Sub InsertNew(ins As Variant, insRow As Variant)

    'Declare function variables
    Dim i As Integer

    'Loop through each insert statement in passthrough array
    For i = 0 To UBound(ins)

        'Execute insert statement & return primary key, customer and program ID
        rst.Open "INSERT INTO UL_Programs " _
            & "OUTPUT inserted.PRIMARY_KEY AS PKEY, " _
            & "inserted.CUSTOMER_ID AS CID, " _
            & "inserted.PROGRAM_ID AS PID " _
            & "VALUES(" & ins(i) & ")", cnn

        'Update Excel file (programs) with primary key, customer and program ID
        Cells(insRow(i), oPrgms.ColIndex("PRIMARY_KEY") + 1).value = _
            rst.Fields("PKEY").value
        Cells(insRow(i), oPrgms.ColIndex("CUSTOMER_ID") + 1).value = _
            rst.Fields("CID").value
        Cells(insRow(i), oPrgms.ColIndex("PROGRAM_ID") + 1).value = _
            rst.Fields("PID").value

        'Close recordset
        rst.Close
    Next
End Sub


Sub Fake_Save()

    Dim Programs As New Scripting.Dictionary
    Dim var     As Variant
    Dim varChange As Variant

    Set Programs = dctPrograms(False)
    var = GetXL("Programs$")

    varChange = Compare(Programs, var)

    If varChange(0)(0) <> 0 Or varChange(1)(0) <> 0 Then

        cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

        If varChange(0)(0) <> 0 Then UploadChanges (varChange(0))
        If varChange(1)(0) <> 0 Then RetVal = InsertNew(varChange(1), varChange(2))

        cnn.Close
        Set cnn = Nothing
    End If
End Sub
