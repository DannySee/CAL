VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "clsDevLoads"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private dictCol As New Scripting.Dictionary
Private dictType As New Scripting.Dictionary
Private dictIndex As New Scripting.Dictionary
Private strSheet As String

Public Property Get Name() As String

    Set Name = strSheet

End Property

Public Property Get Cols() As Scripting.Dictionary

    Set Cols = dictCol

End Property

Public Property Get ColIndex() As Scripting.Dictionary

    Set ColIndex = dictIndex

End Property

Private Sub Class_Initialize()

    strSheet = "Deviation Loads"

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
