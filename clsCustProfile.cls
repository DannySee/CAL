VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "clsCustProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'Declare private class constants
Private Const strSheet As String = "Customer Profile"
Private Const strDb As String = "UL_Customer_Profile"

'Declare private class variables
Private dictCol As New Scripting.Dictionary
Private dictType As New Scripting.Dictionary
Private dictIndex As New Scripting.Dictionary


'*******************************************************************************
'Property to return class sheet name constant (Customer Profile)
'*******************************************************************************
Public Property Get Name() As String

    'Return class sheet name
    Name = strSheet
End Property


'*******************************************************************************
'Property to return class database table constant (UL_Customer_Profile)
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
'Initialize sub. Create dictionaries with column index & name keys
'*******************************************************************************
Private Sub Class_Initialize()

    'Assemble dictionary of field indeces(key) and names(item)
    With dictCol
        .Add 0, "PRIMARY_KEY"
        .Add 1, "CUSTOMER_ID"
        .Add 2, "CUSTOMER"
        .Add 3, "ALT_NAME"
        .Add 4, "PACKET"
        .Add 5, "PRICE_RULE"
        .Add 6, "NID"
        .Add 7, "MASTER_PRN"
        .Add 8, "PRICING_PRN"
        .Add 9, "GROUP_NAME"
        .Add 10, "VPNA"
        .Add 11, "NAM"
        .Add 12, "CUST_CONTACT"
        .Add 13, "NOTES"
    End With

    'Assemble dictionary of field names(key) and indeces(item)
    With dictIndex
        .Add "PRIMARY_KEY", 0
        .Add "CUSTOMER_ID", 1
        .Add "CUSTOMER", 2
        .Add "ALT_NAME", 3
        .Add "PACKET", 4
        .Add "PRICE_RULE", 5
        .Add "NID", 6
        .Add "MASTER_PRN", 7
        .Add "PRICING_PRN", 8
        .Add "GROUP_NAME", 9
        .Add "VPNA", 10
        .Add "NAM", 11
        .Add "CUST_CONTACT", 12
        .Add "NOTES", 13
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
'Returns a static dictionary of Customer Profile data (Key = Primary Key, Value
'= array of fields). Boolean parameter indicates if the dictionary needs to be
'updated before it is returned.
'*******************************************************************************
Function GetDct(blUpdate As Boolean) As Scripting.Dictionary

    'declare static dictionary to hold program data
    Static dct As New Scripting.Dictionary

    'Update dictionary before its returned (if indicated by passthrough boolean)
    If blUpdate Then Set dct = RefreshDct(dct)

    'Return static dictionary
    Set GetDct = dct
End Function
