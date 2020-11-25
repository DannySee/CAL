Attribute VB_Name = "Data_Maintenance"


'*******************************************************************************
'Returns a static dicitonary of Program data (Key = Program ID, Value = array of
'each field). Boolean operator indicates if the dictionary needs to be updated
'before it is returned.
'*******************************************************************************
Function GetDict(blUpdate As Boolean) As Scripting.Dictionary

    'declare static dictionary to hold program data
    Static dctPrograms As New Scripting.Dictionary

    'Update dictionary before its returned (if indicated by passthrough boolean)
    If blUpdate = True Then Set dctPrograms = UpdateDict(dctPrograms)

    'Return static dictionary
    Set GetDict = dctPrograms
End Function


'*******************************************************************************
'Returns dictionary dictionary of program data (Key = Program_ID, Value = Array
'of each field). Meant to update the static the static dictionary that is
'passed through.
'*******************************************************************************
Function UpdateDict(dctPrograms As Scripting.Dictionary) As Scripting.Dictionary

    'Declare function variables
    Dim iRow, iCol As Integer
    Dim iCustCol As Integer
    Dim arr As Variant
    Dim var As Variant

    'Save customer column index
    iCustCol = oPrgms.ColIndex("CUSTOMER_ID")

    'Clear values items
    dctPrograms.RemoveAll

    'Save multidimensional array of program data
    var = GetXL("Programs$")

    'Create an empty array with an index for each program field'
    ReDim arr(UBound(var, 1))

    'Loop through each row of program data
    For iRow = 0 To UBound(var, 2)

        'Loop through each column of program data & add element to array
        For iCol = 0 To UBound(var, 1)
            arr(iCol) = var(iCol, iRow)
        Next

        'Add line to dictionary with program ID as key & row fields as value
        dctPrograms(var(iCustCol, iRow)) = arr
    Next

    'Return dictionary of program data
    Set UpdateDict = dctPrograms
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
    var = rst.GetRows()

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
Function Compare(old As Scripting.Dictionary, upd As Variant) As Variant

    'Declare function variables
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iCIDCol  As integer
    Dim iPIDCol As Integer
    Dim iSDateCol As Integer
    Dim iCustCol As Integer
    Dim iPKeyCol As Integer
    Dim pID As String
    Dim strInsert As String
    Dim strAssemble As String
    Dim strUpdate As String
    Dim strVal As String
    Dim strInsRows As String

    'Get column index for pertinant fields
    iCIDCol = oPrgms.ColIndex("CUSTOMER_ID")
    iCustCol = oPrgms.ColIndex("CUSTOMER")
    iPIDCol = oPrgms.ColIndex("PROGRAM_ID")
    iSDateCol = oPrgms.ColIndex("START_DATE")
    iPKeyCol = oPrgms.ColIndex("PRIMARY_KEY")

    'Loop through rows of current program data
    For iRow = 0 To UBound(upd, 2)

        'If row is new with at least one field filled out (customer name)
        If IsNull(upd(iCIDCol, iRow)) Then
            If upd(iCustCol, iRow) <> "" Then

                'Save SQL insert string (including each element of programs tab)
                strInsert = Append(strInsert, "|") & GetInsertString(upd, iRow)

                'Save array of insert rows
                strInsRows = Append(iRow + 2, "|")
            End If
        Else

            'Get current line program ID (used for SQL string)
            pID = upd(iPIDCol, iRow)

            'Set temp update SQL string to nothing
            strAssemble = ""

            'Loop through columns of current program data
            For iCol = 0 To UBound(upd, 1)

                'If current data does not match static dictionary
                If (old(pID)(iCol) <> upd(iCol, iRow)) Or _
                    (IsNull(old(pID)(iCol)) <> IsNull(upd(iCol, iRow))) Then

                    'Get validated updated value
                    strVal = Validate(upd(iCol, iRow), iCol, iRow)

                    'If update is datetime and entry is valid
                    If strVal <> "'DateErr'" Then

                        'Add entry to udate SQL string
                        strAssemble = Append(strAssemble, ",") _
                            & oPrgms.Cols(iCol) & " = " & strVal
                    End If
                End If
            Next

            'If SQL update string contains start date change
            If InStr(strAssemble, "START_DATE") <> 0 Then

                    'Save SQL update string to end date of previous record
                    strUpdate = Append(strUpdate, "|") & "END_DATE = '" _
                        & upd(iSDateCol, iRow) - 1 & "' " _
                        & "WHERE PRIMARY_KEY = " _
                        & upd(iPKeyCol, iRow)

                    'Save SQL insert string
                    strInsert = Append(strInsert, "|") _
                        & GetInsertString(upd, iRow)

                    'Save an array of insert rows
                    strInsRows = Append(iRow + 2, "|")

            'If SQL update string does not contain start date change
            ElseIf strAssemble <> "" Then

                'Save SQL update string
                strUpdate = Append(strUpdate, "|") & strAssemble _
                    & " WHERE PRIMARY_KEY = " & upd(iPKeyCol, iRow)
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
'Return value which has been validated for its SQL field datatype. This Function
'is meant to assist in assembling SQL update/insert strings. Passthrough
'variables are the string value to be validated and its origin row/column.
'Column/row index is from multidimensional array, not excel coordinates.
'*******************************************************************************
Function Validate(val As Variant, iCol As Integer, iRow As Integer) As String

    'Declare function variables
    Dim sep As String
    Dim iSDateCol As integer
    Dim iEDateCol As integer
    Dim iVendCol As integer

    'Get column index for pertinant fields
    iSDateCol = oPrgms.ColIndex("START_DATE")
    iEDateCol = oPrgms.ColIndex("END_DATE")
    iVendCol = oPrgms.ColIndex("VENDOR_NUM")

    'Get string with datatype delimiters (quotes for text, nothing for number)
    sep = oPrgms.ColType(iCol)

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

    'If passthrough value contains an apostrophe
    ElseIf InStr(val, "'") <> 0 Then

        'Remove apostrophe and save value
        val = Replace(val, "'", "")
    End If

    'Return validated value wrapped in datatype appropriate delimiter
    Validate = sep & val & sep
End Function


'*******************************************************************************
'Returns a concatenated string and separator.
'*******************************************************************************
Function Append(val As String, sep As String) As String

    'If value is blank
    If val = "" Then

        'Return blank variable
        Append = ""
    Else

        'Return concatenated value and separator
        Append = val & sep & " "
    End If
End Function


'*******************************************************************************
'Returns a SQL insert statement. Passthrough variables are a multidimensional
'array with program data and the row index to be inserted.
'*******************************************************************************
Function GetInsertString(var As Variant, iRow As Integer) As String

    'Declare function variables
    Dim iCol As Integer
    Dim iCIDCol As Integer
    Dim iPIDCol As integer
    Dim iDABCol As integer
    Dim strRow As String
    Dim strVal As String

    'Get column index for pertinant fields
    iCIDCol = oPrgms.ColIndex("CUSTOMER_ID")
    iPIDCol = oPrgms.ColIndex("PROGRAM_ID")
    iDABCol = oPrgms.ColIndex("DAB")

    'If array does not contain customer ID data
    If IsNull(var(iCIDCol, iRow)) Then

        'Establish conection to SQL server
        cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1" _
            & ";DATABASE=Pricing_Agreements;"

        'Query customer and program IDs from customer name
        rst.Open "SELECT TOP 1 CUSTOMER_ID AS CID, " _
            & "MAX(CAST(RIGHT(PROGRAM_ID, " _
            & "CHARINDEX('-', REVERSE(PROGRAM_ID)) - 1) AS INT)) + 1 AS PID " _
            & "FROM UL_Programs WHERE CUSTOMER = '" _
            & var(oPrgms.ColIndex("CUSTOMER"), iRow) & "' " _
            & "GROUP BY CUSTOMER_ID", cnn

        'Assemble first 3 fields of SQL insert string
        strRow = rst.Fields("CID").value & ",'" & rst.Fields("CID").value _
            & "-" & rst.Fields("PID").value & "'"

        'Close and free objects
        rst.Close
        cnn.Close
        Set rst = Nothing
        Set cnn = Nothing
    Else

        'Assemble first 3 fields of SQL insert string
        strRow = var(iCIDCol, iRow) & ",'" _
            & var(iPIDCol, iRow) & "'"
    End If

    'Loop through each column of passthrough array
    For iCol = iDABCol To UBound(var, 1)

        'Assemble SQL insert string
        strVal = Validate(var(iCol, iRow), iCol, iRow)

        'Assemble array of insert rows
        strRow = strRow & "," & strVal
    Next

    'Return SQL insert string
    GetInsertString = strRow
End Function


'*******************************************************************************
'Returns no value - executes SQL update statement to CAL database. Passthrough
'variable is an array of update statements, one element per statement.
'*******************************************************************************
Function UploadChanges(upd As Variant)

    'Declare function variables
    Dim i As Integer

    'Loop through each update statement in passthrough array
    For i = 0 To UBound(upd)

        'Execute update statement
        cnn.Execute "UPDATE UL_Programs SET " & upd(i)
    Next
End Function


'*******************************************************************************
'Returns no value - execute SQL insert statement to CAL database. passthrough
'variables are an array of insert statements (one element per statement) and
'an array with the corresponding excel row numbers.
'*******************************************************************************
Function InsertNew(ins As Variant, insRow As Variant)

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
End Function







Sub Fake_Refresh()

    Dim Programs As New Scripting.Dictionary

    Set Programs = GetDict(True)

End Sub

Sub Fake_Save()

    Dim Programs As New Scripting.Dictionary
    Dim var     As Variant
    Dim varChange As Variant

    Set Programs = GetDict(False)
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
