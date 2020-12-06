VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "clsDevLoads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'Declare private class constants
Private Const strSheet As String = "Deviation Loads"
Private Const strDb As String = "UL_Deviation_Loads"
Private Const iPKey As Integer = 0

'Declare private class variables
Private dictCol As New Scripting.Dictionary
Private dictType As New Scripting.Dictionary
Private dictIndex As New Scripting.Dictionary
Private iRow As Integer
Private iCol As Integer


'*******************************************************************************
'Property to return class sheet name constant (Deviation Loads)
'*******************************************************************************
Public Property Get Name() As String

    'Return class sheet name
    Name = strSheet
End Property


'*******************************************************************************
'Property to return class database table name constant (UL_Deviation_Loads)
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
'*******************************************************************************\
Public Property Get ColIndex() As Scripting.Dictionary

    'Return class database field dictionary
    Set ColIndex = dictIndex
End Property


'*******************************************************************************
'Initialize sub. Create dictionaries with column index & name keys
'*******************************************************************************
Private Sub Class_Initialize()

    'Assemble dictionary of field indeces(key) and names(item)
    With dictCol
        .Add 0, "PRIMARY_KEY"
        .Add 1, "CUSTOMER_ID"
        .Add 2, "CUSTOMER"
        .Add 3, "PROGRAM"
        .Add 4, "OWNER"
        .Add 5, "DATE"
        .Add 6, "SUPPORT_REQUEST"
        .Add 7, "GRP"
    End With

    'Assemble dictionary of field names(key) and indeces(item)
    With dictIndex
        .Add "PRIMARY_KEY", 0
        .Add "CUSTOMER_ID", 1
        .Add "CUSTOMER_NAME", 2
        .Add "PROGRAM", 3
        .Add "OWNER", 4
        .Add "DATE", 5
        .Add "SUPPORT_REQUEST", 6
        .Add "GRP", 7
    End With
End Sub


'*******************************************************************************
'Returns updated dictionary of Excel data (Key = Primary_Key, Value = Array
'of fields). Meant to update the passthrough dictionary with static dictionary.
'*******************************************************************************
Function RefreshDct() As Scripting.Dictionary

    'Declare function variables
    Dim dct As New Scripting.Dictionary
    Dim arr As Variant
    Dim var As Variant

    'Save multidimensional array of program data
    var = GetXL(strSheet)

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
'Returns a static dictionary of Deviation Loads (Key = Primary Key, Value =
'array of fields). Boolean parameter indicates if the dictionary needs to be
'updated before it is returned.
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
'Return value which has been validated for its SQL field datatype. This Function
'is meant to assist in assembling SQL update/insert strings to the UL_Programs
'database. Passthrough variables are the string value to be validated and its
'origin row/column. Boolean parameter indicates if the data type needs to be
'validated. Column/row index is from multidimensional array, not Excel
'coordinates.
'*******************************************************************************
Function Validate(val As Variant, iCol As Integer, iRow As Integer) As String

    'Declare function variables
    Dim sep As String

    'Set string delimiter to SQL syntax single quotes
    sep = "'"

    'Return validated value wrapped in datatype appropriate delimiter
    Validate = sep & Replace(val, "'", "") & sep
End Function


'*******************************************************************************
'Returns an SQL insert statement. Function is Specific to Customer Profile tab.
'Gutter data is Primary Key and Cusotmer ID. Guts data is all other Customer
'Profile fields. Parameters are multidimensional array of Customer Profile tab
'and focus row index.
'*******************************************************************************
Function GetIns(var As Variant, iRow As integer) As String

    'Declare function variables
    Dim strGutter As String
    Dim strGuts As String

    'If array does not contain customer ID data
    If IsNull(var(iPKey, iRow)) Then

        'Establish conection to SQL server
        cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1" _
            & ";DATABASE=Pricing_Agreements;"

        'Query customer and program IDs from customer name
        rst.Open "SELECT TOP 1 CUSTOMER_ID AS CID " _
            & "FROM " & strDb & " WHERE CUSTOMER = '" _
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
    strGuts = AppendRow(var, iRow, iCst)

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

            'Add entry to udate SQL string
            strUpd = Append(strUpd, ",", _
                dictCol(iCol) & " = " & strVal)
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

            'Add PRIMARY_KEY to to return string (comma separator)
            Append(strDel, ",", key)
        End if
    Next

    'Return string of deletions
    GetDel = strDel
End Function


'*******************************************************************************
'Return multidimensional array of worksheet data that was updated since the
'static dictionary was initialized. The first index of the return array contains
'Deviation Loads elements to be updated. The second index contains program
'elements to be inserted. Parameters are the static dictionary with historical
'Deviation Load data and a multidimensional array of current Deviation Load data
'*******************************************************************************
Function GetChanges() As Scripting.Dictionary

    'Declare function variables
    Dim dbChanges As Scripting.Dictionary
    Dim old As New Scripting.Dictionary
    Dim upd As Variant
    Dim pKey As String
    Dim strIns As String
    Dim strUpd As String
    Dim strDel As String
    Dim strInsRows As String

    'Get static dictionary data (historical data for comparison)
    Set old = GetSaveData(False)

    'Get new dictionary data (new data for comparison)
    Set upd = GetXL(strSheet)

    'Loop through rows of current program data
    For iRow = 0 To UBound(upd, 2)

        'If row is new with at least one field filled out (customer name)
        If IsNull(upd(iPKey, iRow)) And Not IsNull(upd(iCst, iRow)) Then

            'Save SQL insert string (including each element of Excel tab)
            strIns = Append(strIns, "|", GetIns(upd, iRow))

            'Save array of insert excel row
            strInsRows = Append(strInsRows, "|", iRow + 2)

        'If row is not new
        ElseIf Not IsNull(upd(iPKey, iRow))

            'Get SQL update string of any worksheet updates
            strUpd = Append(strUpd, "|", GetUpd(old, upd, iRow)) _
                & " WHERE PRIMARY_KEY = " & upd(iPKey, iRow))
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

    'Return dictionary with update/insert/delete SQL strings
    Set GetChanges = dbChanges
End Function
