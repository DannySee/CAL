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
Private Const strSheet = "Customer Profile"
Private Const strDb = "UL_Customer_Profile"

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
