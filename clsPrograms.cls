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
Private Const strSheet = "Programs"
Private Cosnt strDb = "UL_Programs"

'Declare private class variables
Private dictCol As New Scripting.Dictionary
Private dictType As New Scripting.Dictionary
Private dictIndex As New Scripting.Dictionary


'*******************************************************************************
'Property to return class sheet name constant (Programs)
'*******************************************************************************
Public Property Get Name() As String

    'Return class sheet name
    Name = strSheet
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
