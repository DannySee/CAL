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
Private Const strSht As String = "Programs"
Private Const strDb As String = "UL_Programs"
Private Const strDbx As String = "Deleted_Programs"
Private Const iPKey As Integer = 0
Private Const iCID As Integer = 1
Private Const iPID As Integer = 2
Private Const iDAB As Integer = 3
Private Const iTime As Integer = 5
Private Const iTier As Integer = 6
Private Const iCst As Integer = 7
Private Const iSDte As Integer = 9
Private Const iEDte As Integer = 10
Private Const VAType As Integer = 13
Private Const iVend As Integer = 14
Private Const iBB As Integer = 15
Private Const iCost As Integer = 16
Private Const iCAType As Integer = 17
Private Const iRebate As Integer = 18
Private Const iApprop As Integer = 20
Private Const strCstHeaders As Variant = Array("PROGRAM DESCRIPTION", _
    "START DATE","END DATE","LEAD VA","LEAD CA","VA TYPE","VENDOR", _
    "BB FORMAT","COST BASIS","CA TYPE","REBATE BASIS","PRE APPROVAL", _
    "APPROP NAME","PRN/GRP","PACKET","PACKET DL","COMMENTS","","","","","")
Private Const strCstFlds As String = _
    & "ROGRAM_DESCRIPTION, " _
    & "START_DATE, " _
    & "END_DATE, " _
    & "LEAD_VA, " _
    & "LEAD_CA, " _
    & "VEND_AGMT_TYPE, " _
    & "VENDOR_NUM, " _
    & "BILLBACK_FORMAT, " _
    & "COST_BASIS, " _
    & "CUST_AGMT_TYPE, " _
    & "REBATE_BASIS, " _
    & "PRE_APPROVAL, " _
    & "APPROP_NAME, " _
    & "PRN_GRP, " _
    & "PACKET, " _
    & "PACKET_DL, " _
    & "COMMENTS"

'Declare private class variables
Private dictCol As New Scripting.Dictionary
Private dictType As New Scripting.Dictionary
Private dictIndex As New Scripting.Dictionary
Private iCol As Integer
Private iRow As Integer


'*******************************************************************************
'Property to return customer friendly headers
'*******************************************************************************
Public Property Get CstHeaders() As String

    'Return class sheet name
    CstHeaders = strCstHeaders
End Property


'*******************************************************************************
'Property to return customer friendly Db field names
'*******************************************************************************
Public Property Get CstFlds() As String

    'Return class sheet name
    CstFlds = strCstFlds
End Property


'*******************************************************************************
'Property to return class sheet name constant (Programs)
'*******************************************************************************
Public Property Get Sht() As String

    'Return class sheet name
    Sht = strSht
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
'Returns a static dictionary of Program data (Key = Primary Key, Value = array
'of fields). Boolean parameter indicates if the dictionary needs to be updated
'before it is returned.
'*******************************************************************************
Function GetSaveData(blUpdate As Boolean) As Scripting.Dictionary

    'declare static dictionary to hold program data
    Static dct As New Scripting.Dictionary

    'Update dictionary before its returned (if indicated by passthrough boolean)
    If blUpdate Then Set dct = RefreshDct(strSht)

    'Return static dictionary
    Set GetSaveData = dct
End Function


'*******************************************************************************
'Add data validaiton to sheet. Include appropriate drop down list for all
'restricted fields. BB Format will always include a blank list (field is
'deactivated).
'*******************************************************************************
Sub AddDataValidation()

    'Declare sub variables
    Dim rngYN As Range
    Dim rngDrp As Range
    Dim varRng As Variant
    Dim iDropLR As Integer

    'Get last row
    iLRow = LastRow(strSht)

    'Activate programs tab
    With Sheets(strSht)

        'Remove data validation from sheet
        .Cells.Validation.Delete

        'Set range for fields which will have Y/N drop down
        Set rngYN = .Range(.Cells(2, iDAB + 1), .Cells(iLRow, iTime + 1))

        'Create an array of pertinant Excel (Programs) ranges
        varRng = Array( _
            .Range(.Cells(2, iTier), .Cells(iLRow, iTier)), _
            .Range(.Cells(2, iVAType), .Cells(iLRow, iVAType)), _
            .Range(.Cells(2, iCost), .Cells(iLRow, iCost)), _
            .Range(.Cells(2, iBB), .Cells(iLRow, iBB)), _
            .Range(.Cells(2, iCAType), .Cells(iLRow, iCAType)), _
            .Range(.Cells(2, iRebate), .Cells(iLRow, iRebate)), _
            .Range(.Cells(2, iApprop), .Cells(iLRow, iApprop)), _
            .Range(.Cells(2, iCst), .Cells(iLRow, iCst)))
    End With

    'Add Y/N drop down list to first three editable Fields
    rngYN.Validation.Add xlValidateList, Formula1:="Y,N"

    'Activate DropDowns tab
    With Sheets("DropDowns")

        'Loop through each column of DropDowns tab
        For i = 0 To UBound(varRng)

            'Get last row of dropdown fields (custom columns)
            iDropLR = .Cells(.Rows.Count, i + 1).End(xlUp).Row

            'Save dropdown Range
            Set rngDrp = .Range(.Cells(1, i + 1), .Cells(iDropLR, i + 1))

            'Add data validation to Excel (Programs) range
            varRng(i).Validation.Add xlValidateList, _
                Formula1:="=DropDowns!" & rngDrp.Address
        Next
    End With
End Sub


'*******************************************************************************
'Return value which has been validated for its SQL field datatype. This Function
'is meant to assist in assembling SQL update/insert strings to the UL_Programs
'database. Passthrough variables are the string value to be validated and its
'origin row/column. Column/row index is from multidimensional array, not Excel
'coordinates.
'*******************************************************************************
Function Validate(val As Variant, iCol As Integer, iRow As Integer) As String

    'Declare function variables
    Dim sep As String

    'Get string with datatype delimiters (quotes for text, blank for number)
    sep = dictType(iCol)

    'If passthrough data type is datetime and value is invalid date
    If (iCol = iSDte Or iCol = iEDte) And Not IsDate(val) Then

        'Alert user of invalid date value
        MsgBox "INVALID DATE: Please correct the " & dictCol(iCol) _
            & " field - " & val & " is not a valid date entry"

        'Highlight (select) row with invalid date value to alert user
        Rows(iRow + 2).EntireRow.Select

        'Return invalid date error message
        val = "DateErr"

    'If passthrough datatype is number and value is empty
    ElseIf iCol = iVend And IsNull(val) Then

        'Return 0 value (string to maintain function type)
        val = "0"
    End if

    'Return validated value wrapped in datatype appropriate delimiter
    Validate = sep & Replace(val, "'", "") & sep
End Function


'*******************************************************************************
'Returns an SQL insert statement. Gutter data is Primary Key, Customer ID and
'Program ID. Guts data is all other program fields. Parameters are
'multidimensional array of programs tab and focus row index.
'*******************************************************************************
Function GetIns(var As Variant, iRow As integer) As String

    'Declare function variables
    Dim strGutter As String
    Dim strGuts As String

    'If array does not contain customer ID data
    If IsNull(var(iPKey, iRow)) Then

        'Assemble first 3 fields of SQL insert string
        strGutter = Pull.GetCstPgmID(var(iCst, iRow))
    Else

        'Assemble first 3 fields of SQL insert string
        strGutter = var(iCID, iRow) & ",'" & var(iPID, iRow) & "'"
    End If

    'Get concatenated string of all fields in SQL sytnax
    strGuts = AppendRow(var, iRow, iDAB)

    'Return Programs gutter
    GetIns = Append(strGutter, ",", strGuts)
End Function


'*******************************************************************************
'Returns an SQL update statement. Gutter data is Primary Key, Customer ID and
'Program ID. Guts data is all other program fields. Parameters are
'multidimensional array of programs tab and focus row index.
'*******************************************************************************
Function GetUpd(old As Scripting.Dictionary, upd As Variant, _
    iRow As Integer) As String

    'Delcare function variables
    Dim pKey As String
    Dim strVal As String
    Dim strUpd As String

    'Get current line program ID (used for SQL string)
    pKey = upd(iPKey, iRow)

    'Loop through columns of current program data
    For iCol = 0 To UBound(upd, 1)

        'If current data does not match static dictionary
        If (old(pKey)(iCol) <> upd(iCol, iRow)) Or _
            (IsNull(old(pKey)(iCol)) <> IsNull(upd(iCol, iRow))) Then

            'Get validated updated value
            strVal = Validate(upd(iCol, iRow), iCol, iRow)

            'If update is datetime and entry is valid
            If strVal <> "'DateErr'" Then

                'Add entry to udate SQL string
                strUpd = Append(strUpd, ",", dictCol(iCol) & " = " & strVal)
            End If
        End If
    Next

    'Return update strings
    GetUpd = strUpd
End Function


'*******************************************************************************
'Binary search algorithm returns a string of PROGRAM_IDs (for program tab) &
'PRIMARY_KEYs (for customer profile & deviation loads) which are to be deleted
'from their respective data tables.
'*******************************************************************************
Function GetDel(old As Scripting.Dictionary upd As Variant) As string

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

            'Add PROGRAM_ID to return string (comma separator)
            Append(strDel, ",", old(Key)(2))
        End if
    Next

    'Return string of deletions
    GetDel = strDel
End Function


'*******************************************************************************
'Return multidimensional array of program data that was updated since the Static
'dictionary was initialized. The first index of the return array contains
'program elements to be updated. The second index contains program elements to
'be inserted. Parameters are the static dictionary with historical program data
'and a multidimensional array of current program data.
'*******************************************************************************
Function GetChanges() As Scripting.Dictionary

    'Declare function variables
    Dim dbChanges As Scripting.Dictionary
    Dim old As New Scripting.Dictionary
    Dim upd As Variant
    Dim strIns As String
    Dim strUpd As String
    Dim strDel As String
    Dim strVal As String
    Dim strInsRows As String

    'Get static dictionary data (historical data for comparison)
    Set old = GetSaveData(False)

    'Get new dictionary data (new data for comparison)
    Set upd = Pull.GetXL(strSht)

    'Loop through rows of current program data (will retrieve upd/ins strings)
    For iRow = 0 To UBound(upd, 2)

        'If row is new with at least one field filled out (customer name)
        If IsNull(upd(iPKey, iRow)) And Not IsNull(upd(iCst, iRow)) Then

            'Save SQL insert string (including each element of programs tab)
            strIns = Append(strIns, "|", GetIns(upd, iRow))

            'Save array of Excel insert rows
            strInsRows = Append(strInsRows, "|", iRow + 2)

        'If row is not new
        ElseIf Not IsNull(upd(iPKey, iRow))

            'Get SQL update string of any worksheet updates
            strVal = GetUpd(old, upd, iRow)

            'If SQL update string contains start date change
            If InStr(strVal, "START_DATE") <> 0 Then

                'Save SQL update string to end date of previous record
                strUpd = Append(strUpd, "|", "END_DATE = '" _
                    & upd(iSDte, iRow) - 1 & "' " _
                    & "WHERE PRIMARY_KEY = " _
                    & upd(iPKey, iRow))

                'Save SQL insert string
                strIns = Append(strIns, "|", & GetIns(upd, iRow))

                'Save an array of Excel insert rows
                strInsRows = Append(strinsRows, "|", iRow + 2)

            'If SQL update string does not contain start date change
            ElseIf strVal <> "" Then

                'Save SQL update string
                strUpd = Append(strUpd, "|", strVal _
                    & " WHERE PRIMARY_KEY = " & upd(iPKey, iRow))
            End If
        End If
    Next

    'Get SQL delete string of any worksheet deletions
    strDel = GetDel(old, upd)

    'If there are any updated/deleted records, add them to change dictionary
    If strUpd <> "" Then dbChanges("UPDATE") = Split(strUpd, "|")
    If strDel <> "" Then dbChanges("DELETE") = strDel

    'If there are any new records add them (and their excel index) to dictionary
    If strIns <> "" Then
        dbChanges("INSERT") = Split(strIns, "|")
        dbChanges("ROWS") = Split(strInsRows, "|")
    End If

    'Return multidimensional array w/ update(0) and insert(1) SQL string arrays
    Set GetChanges = dbChanges
End Function


'*******************************************************************************
'Execute SQL insert statement to CAL database. passthrough variables are an
'array of insert statements (one element per statement) and an array with the
'corresponding excel row numbers.
'*******************************************************************************
Sub DeleteRecords(strDel As String)

    'Declare sub variables
    Dim varDel As Variant
    Dim strIns As String

    'Get multidimensional array of deleted prgram elements
    varDel = Push.GetDeleted(strDel, strDb)

    'Get SQL insert string fro archive tables & add newtwork ID
    strIns = GetIns(varDel, 0) & ",'" & Environ("Username") & "'"

    'Insert deleted record into archive table
    Push.InsertDeleted(strIns, strDbx)
End Sub


'*******************************************************************************
'Execute SQL insert statement to CAL database. passthrough variables are an
'array of insert statements (one element per statement) and an array with the
'corresponding excel row numbers.
'*******************************************************************************
Sub InsertNew(ins As Variant, insRow As Variant)

    'Declare function variables
    Dim retVal As ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer

    'Loop through each insert statement in passthrough array
    For i = 0 To UBound(ins)

        'Set SQL string insert statement
        strSQL = "INSERT INTO " & strDb & " " _
            & "OUTPUT inserted.PRIMARY_KEY AS PKEY, " _
            & "inserted.CUSTOMER_ID AS CID, " _
            & "inserted.PROGRAM_ID AS PID " _
            & "VALUES(" & ins(i) & ")"

        'Insert elements and return new customer ID and Primary Key
        Set retVal = Push.Insert(strSQL)

        'Update Excel file with primary key, customer ID & program ID
        Cells(insRow(i), iPkey).value = retVal.Fields("PKEY").value
        Cells(insRow(i), iCID).value = retVal.Fields("CID").value
        Cells(insRow(i), iPID).value = retVal.Fields("PID").value
    Next
End Sub


'*******************************************************************************
'Upload/inserts new worksheet records to the SQL server. Refresh static
'dictionary after upload is complete.
'*******************************************************************************
Sub Push()

    'Declare sub variables
    Dim dbChanges As New Scripting.Dictionary

    'Retrieve workbook changes since last save
    Set dbChanges = GetChanges

    'If updates were made to worksheet fields, push updates to server
    If dbChanges.Exists("UPDATE") Then Push.Update(dbChanges("UPDATE"), strDb)

    'If updates were made to worksheet fields, push updates to server
    If dbChanges.Exists("DELETE") Then DeleteRecords(dbChanges("DELETE"))

    'If new lines were added to worksheet, push additionas to server
    If dbChanges.Exists("INSERT") Then InsertNew(dbChanges("INSERT"), _
        dbChanges("ROWS"))

    'Refresh static dictionary to save new worksheet data
    Set dbChanges = GetSaveData(True)
End Sub
