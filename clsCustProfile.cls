VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "clsCustProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private dictCol As New Scripting.Dictionary
Private dictType As New Scripting.Dictionary
Private dictIndex As New Scripting.Dictionary
Private strFields As String

Public Property Get Cols() As Scripting.Dictionary

    Set Cols = dictCol

End Property

Public Property Get ColIndex() As Scripting.Dictionary

    Set ColIndex = dictIndex

End Property

Private Sub Class_Initialize()

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
