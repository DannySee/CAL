Attribute VB_Name = "Data_Maintenance"


'*******************************************************************************
'Returns a static dictionary of Program data (Key = Primary Key, Value = array
'of fields). Boolean parameter indicates if the dictionary needs to be updated
'before it is returned.
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
'Returns a static dictionary of Customer Profile data (Key = Primary Key, Value
'= array of fields). Boolean parameter indicates if the dictionary needs to be
'updated before it is returned.
'*******************************************************************************
Function dctCstProfile(blUpdate As Boolean) As Scripting.Dictionary

    'declare static dictionary to hold program data
    Static dctCst As New Scripting.Dictionary

    'Update dictionary before its returned (if indicated by passthrough boolean)
    If blUpdate Then Set dctCst = UpdateDct(dctCst, "Customer Profile")

    'Return static dictionary
    Set dctCstProfile = dctCst
End Function


'*******************************************************************************
'Returns a static dictionary of Deviation Loads (Key = Primary Key, Value =
'array of fields). Boolean parameter indicates if the dictionary needs to be
'updated before it is returned.
'*******************************************************************************
Function dctDevLds(blUpdate As Boolean) As Scripting.Dictionary

    'declare static dictionary to hold program data
    Static dctDev As New Scripting.Dictionary

    'Update dictionary before its returned (if indicated by passthrough boolean)
    If blUpdate Then Set dctDev = UpdateDct(dctDev, "Deviation Loads")

    'Return static dictionary
    Set dctDevLds = dctDev
End Function


'*******************************************************************************
'Returns updated dictionary of Excel data (Key = Primary_Key, Value = Array
'of fields). Meant to update the passthrough static dictionary from parameter
'sheet.
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
    var = GetXL(strSht)

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
'Return multidimensional array of Excel file. Parameter indicates which tab's
'data should be returned.
'*******************************************************************************
Function GetXL(strSheet As String) As Variant

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
'Return multidimensional array of program data that was updated since the Static
'dictionary was initialized. The first index of the return array contains
'program elements to be updated. The second index contains program elements to
'be inserted. Parameters are the static dictionary with historical program data
'and a multidimensional array of current program data.
'*******************************************************************************
Function ComparePrgms(old As Scripting.Dictionary, upd As Variant) As Variant

    'Declare function variables
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iSDte As Integer
    Dim iCst As Integer
    Dim iPKey As Integer
    Dim pKey As String
    Dim strInsert As String
    Dim strAssemble As String
    Dim strUpdate As String
    Dim strVal As String
    Dim strInsRows As String

    'Get fields index for pertinant data
    iPKey = oPrgms.ColIndex("PRIMARY_KEY")
    iCst = oPrgms.ColIndex("CUSTOMER")
    iSDte = oPrgms.ColIndex("START_DATE")

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
    ComparePrgms = Array(Split(strUpdate, "|"), Split(strInsert, "|"), _
        Split(strInsRows, "|"))
End Function


'*******************************************************************************
'Return multidimensional array of Deviation Load data that was updated since the
'Static dictionary was initialized. The first index of the return array contains
'Deviation Loads elements to be updated. The second index contains program
'elements to be inserted. Parameters are the static dictionary with historical
'Deviation Load data and a multidimensional array of current Deviation Load data
'*******************************************************************************
Function CompareCstDev(old As Scripting.Dictionary, upd As Variant, _
    oSht As Object) As Variant

    'Declare function variables
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iCst As Integer
    Dim iPKey As Integer
    Dim pKey As String
    Dim strInsert As String
    Dim strAssemble As String
    Dim strUpdate As String
    Dim strVal As String
    Dim strInsRows As String

    'Get column index for pertinant fields
    iPKey = oSht.ColIndex("PRIMARY_KEY")
    iCst = oSht.ColIndex("CUSTOMER")

    'Loop through rows of current program data
    For iRow = 0 To UBound(upd, 2)

        'If row is new with at least one field filled out (customer name)
        If IsNull(upd(iPKey, iRow)) Then
            If upd(iCst, iRow) <> "" Then

                'Save SQL insert string (including each element of Excel tab)
                If oSht.Name = "Customer Profile" Then
                    strInsert = Append(strInsert, "|", GetCstInsert(upd, iRow))
                ElseIf oSht.Name = "Deviation Loads" Then
                    strInsert = Append(strInsert, "|", GetDevInsert(upd,iRow))
                End If

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

                    'Add entry to udate SQL string
                    strAssemble = Append(strAssemble, ",", _
                        oSht.Cols(iCol) & " = " & strVal)
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
    CompareCstDev = Array(Split(strUpdate, "|"), Split(strInsert, "|"), _
        Split(strInsRows, "|"))
End Function


'*******************************************************************************
'Return value which has been validated for its SQL field datatype. This Function
'is meant to assist in assembling SQL update/insert strings to the UL_Programs
'database. Passthrough variables are the string value to be validated and its
'origin row/column. Boolean parameter indicates if the data type needs to be
'validated. Column/row index is from multidimensional array, not Excel
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
'Returns a concatenated string with separator.
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
'Data is Primary Key, Customer ID and Program ID. Guts data is all other program
'fields. Parameters are multidimensional array of programs tab and
'focus row index.
'*******************************************************************************
Function GetPrgmInsert(var As Variant, iRow As integer) As String

    'Declare function variables
    Dim iPKey As Integer
    Dim iCID As integer
    Dim iPID As integer
    Dim iDAB As integer
    Dim iCst As Integer
    Dim strGutter As String
    Dim strGuts As String

    'Get column index for pertinant fields
    iPKey = oPrgms.ColIndex("PRIMARY_KEY")
    iCID = oPrgms.ColIndex("CUSTOMER_ID")
    iPID = oPrgms.ColIndex("PROGRAM_ID")
    iDAB = oPrgms.ColIndex("DAB")
    iCst = oPrgms.ColIndex("CUSTOMER")

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
            & var(iCst, iRow) & "' " _
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
'Returns an SQL insert statement. Function is Specific to Customer Profile tab.
'Gutter data is Primary Key and Cusotmer ID. Guts data is all other Customer
'Profile fields. Parameters are multidimensional array of Customer Profile tab
'and focus row index.
'*******************************************************************************
Function GetCstDevInsert(var As Variant, iRow As integer) As String

    'Declare function variables
    Dim iPKey As Integer
    Dim iCID As integer
    Dim iCst As integer
    Dim strGutter As String
    Dim strGuts As String

    'Get column index for pertinant fields
    iPKey = oCst.ColIndex("PRIMARY_KEY")
    iCID = oCst.ColIndex("CUSTOMER_ID")
    iCst = oCst.ColIndex("CUSTOMER")

    'If array does not contain customer ID data
    If IsNull(var(iPKey, iRow)) Then

        'Establish conection to SQL server
        cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1" _
            & ";DATABASE=Pricing_Agreements;"

        'Query customer and program IDs from customer name
        rst.Open "SELECT TOP 1 CUSTOMER_ID AS CID " _
            & "FROM UL_Customer_Profile WHERE CUSTOMER = '" _
            & var(iCst, iRow) & "' ", cnn

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
    strGuts = AppendRow(var, iRow, iCst, False)

    'Return Programs gutter
    GetCstInsert = Append(strGutter, ",", strGuts)
End Function


'*******************************************************************************
'Returns an SQL insert statement. Function is Specific to Deviation Loads tab.
'Gutter data is Primary Key and Cusotmer ID. Guts data is all other Deviation
'Loads fields. Parameters are multidimensional array of Deviation Loads tab and
'focus row index.
'*******************************************************************************
Function GetDevInsert(var As Variant, iRow As integer) As String

    'Declare function variables
    Dim iPKey As Integer
    Dim iCID As integer
    Dim iCst As integer
    Dim strGutter As String
    Dim strGuts As String

    'Get column index for pertinant fields
    iPKey = oDev.ColIndex("PRIMARY_KEY")
    iCID = oDev.ColIndex("CUSTOMER_ID")
    iCst = oDev.ColIndex("CUSTOMER_NAME")

    'If array does not contain customer ID data
    If IsNull(var(iPKey, iRow)) Then

        'Establish conection to SQL server
        cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1" _
            & ";DATABASE=Pricing_Agreements;"

        'Query customer and program IDs from customer name
        rst.Open "SELECT TOP 1 CUSTOMER_ID AS CID " _
            & "FROM UL_Deviation_Loads WHERE CUSTOMER_NAME = '" _
            & var(iCst, iRow) & "' ", cnn

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
    strGuts = AppendRow(var, iRow, iCst, False)

    'Return Programs gutter
    GetDevInsert = Append(strGutter, ",", strGuts)
End Function


'*******************************************************************************
'Returns a concatenated string in SQL syntax. Parameters are a multidimensional
'array with Excel data, the row index to be parsed and the starting field index.
'Boolean value is to indicate if validation method requires data type check.
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
'Binary search algorithm returns a string of PROGRAM_IDs (for program tab) &
'PRIMARY_KEYs (for customer profile & deviation loads) which are to be deleted
'from their respective data tables.
'*******************************************************************************
Function GetDelID(old As Scripting.Dictionary upd As Variant) As string

    'Declare function variables
    Dim blFound As Boolean
    Dim strDel As String
    Dim iUpper As Integer
    Dim iLower As Integer
    Dim iMid As Integer

    'Loop through each key (PRIMARY_KEY) in static dictionary
    For Each key in old.keys

        'Set high/low array index and boolean operator to identify matches
        iLower = 0
        iUpper = UBound(upd, 2)
        blFound = False

        'Loop through array while a match has not been found
        Do While blFound = False and iUpper >= iLower

            'Set mid point of array
            iMid = iLower + (iUpper - iLower)/2

            'If key value is greater than the array's middle index value
            If key > upd(0,iMid) Then

                'Set the low search index equal to current middle index + 1
                iLower = iMid + 1

            'If key value is less than the array's middle index value
            ElseIf key < upd(0,iMid) Then

                'Set the high search index equal to current middle index - 1
                iUpper = iMid - 1

            'If key value is equal to the array's middle index value
            Else

                'Identify a match
                blFound =True
            End If
        Loop

        'If a match was not found
        If blFound = False Then

            'If the object sheet is the Programs tab
            If old(key)(2) = "PROGRAM_ID" Then

                'Add PROGRAM_ID to return string (comma separator)
                Append(strDel, ",", old(Key)(2))

            'If the object sheet is the Customer Profile or Deviation Loads tab
            Else

                'Add PRIMARY_KEY to to return string (comma separator)
                Append(strDel, ",", key)
            End If
        End if
    Next

    'Return string of deletions
    GetDelID = strDel
End Function


'*******************************************************************************
'Returns Program ID for new line. Parameter is the Primary Key by which the
'Program ID should be pulled.
'*******************************************************************************
Function GetPID(pKey As Long) As string

    'Declare functiuon variables
    Dim pID As string

    'Pull Program ID from the PROGRAMS table using the Primary Key parameter
    rst.Open "SELECT PROGROM_ID " _
        & "FROM PROGRAMS " _
        & "WHERE PRIMARY_KEY = " & pKey, cnn

    'Save query results
    pID = rst.Fields("PROGRAM_ID").Value

    'Close recordset
    rst.Close

    'Return Program ID
    GetPID = pID
End Function


'*******************************************************************************
'Executes SQL update statement to CAL database. Paramaters are an array of
'update statements (one element per statement) and the SQL table.
'*******************************************************************************
Sub UploadChanges(upd As Variant, strDb As String)

    'Declare function variables
    Dim i As Integer

    'Loop through each update statement in passthrough array
    For i = 0 To UBound(upd)

        'Execute update statement
        cnn.Execute "UPDATE " & strDb " SET " & upd(i)
    Next
End Sub


'*******************************************************************************
'Execute SQL insert statement to CAL database. passthrough variables are an
'array of insert statements (one element per statement) and an array with the
'corresponding excel row numbers.
'*******************************************************************************
Sub InsertNew(ins As Variant, insRow As Variant, strDb As String)

    'Declare function variables
    Dim i As Integer
    Dim iPkey As Integer
    Dim iCID As Integer
    Dim iPID As Integer

    'Get column (Excel) index for gutter columns
    iPkey = 1
    iCID = 2
    iPID = 3

    'Loop through each insert statement in passthrough array
    For i = 0 To UBound(ins)

        'Execute insert statement & return primary key, customer and program ID
        rst.Open "INSERT INTO " & strDb & " " _
            & "OUTPUT inserted.PRIMARY_KEY AS PKEY, " _
            & "inserted.CUSTOMER_ID AS CID, " _
            & "VALUES(" & ins(i) & ")", cnn

        'Update Excel file with primary key & customer ID
        Cells(insRow(i), iPkey).value = rst.Fields("PKEY").value
        Cells(insRow(i), iCID).value = rst.Fields("CID").value

        'Close recordset
        rst.Close

        'Is focus sheet Programs tab and program ID field blank?
        If strDb = "UL_Progrmas" And Cells(insRow(i), iPID).Value = "" Then _
            Cells(insRow(i), iPID).Value = GetPID(Cells(insRow(i), iPKey).Value)
    Next
End Sub


'*******************************************************************************
'Execute SQL insert statement to CAL database. passthrough variables are an
'array of insert statements (one element per statement) and an array with the
'corresponding excel row numbers.
'*******************************************************************************
Sub DeleteRecords(strDel As String, strDb As String)

    Dim strField As string

    If strDb = "UL_Programs" Then
        strField = "PROGRAM_ID"
    Else
        strField = "PRIMARY_KEY"
    End If

    cnn.Execute "DELETE FROM " & strDb _
        & "WHERE " & strField & " IN (" & strDel & ")"

End Sub


'*******************************************************************************
'Upload/inserts new worksheet records to the SQL server. Refresh static
'dictionary after upload is complete.
'*******************************************************************************
Sub Push()

    'Declare sub variables
    Dim dctUpd As New Scripting.Dictionary
    Dim dctDel As New Scripting.Dictionary
    Dim updPrgms As Variant
    Dim updCst  As Variant
    Dim updDev As Variant

    'Retrieve static dictionary with historical database
    Programs = dctPrograms(False)
    cstProfile = dctCstProfile(False)
    devLds = dctDevLds(False)

    'Get multidimensional array of current data
    updPrgms = GetXL(oPrgms.Name))
    updCst = GetXL(oCst.Name))
    updDev = GetXL(oDev.Name))

    'Create dictionary where key = database table and item is array of updates
    With dctUpd
        .Add oPrgms.Db, ComparePrgms(Programs, updPrgms)
        .Add oCst.Db, CompareCstDev(cstProfile, updCst, oCst)
        .Add oDev.Db, CompareCstDev(devLds, updDev, oDev)
    End With

    'Create dictionary where key = database table and item is array of updates
    With dctDel
        .Add oPrgms.Db, GetDelID(Programs, updPrgms)
        .Add oCst.Db, GetDelID(cstProfile, updCst)
        .Add oDev.Db, GetDelID(devLds, updDev)
    End With

    'Establish connection to SQL server
    cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;" _
        & "DATABASE=Pricing_Agreements;"

    'Loop through each key in dictionary
    For Each key In dctUpd

        'If deletions were made, delete records from server
        If dctDel(key) <> "" Then DeleteRecords(dictDel(key), key)

        'If updates were made, push updates to the server
        If dctUpd(key)(0)(0) <> 0 Then UploadChanges(dctUpd(key)(0), key)

        'If new lines were inserted, push inserted lines to the server
        If dctUpd(key)(1)(0) <> 0 Then InsertNew(dctUpd(key)(1), _
            dctUpd(key)(2), key)
    Next

    'Close connection and free objects
    cnn.Close
    Set cnn = Nothing

    'Refresh static dictionaries with new data
    Set Programs = dctPrograms(True)
    Set cstProfile = dctCstProfile(True)
    Set devLds = dctDevLds(True)
End Sub
