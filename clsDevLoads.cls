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
Private Const strSheet = "Deviation Loads"
Private Const strDb = "UL_Deviation_Loads"

'Declare private class variables
Private dictCol As New Scripting.Dictionary
Private dictType As New Scripting.Dictionary
Private dictIndex As New Scripting.Dictionary


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
