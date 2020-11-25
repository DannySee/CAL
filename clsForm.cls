VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "clsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private strPWD As String
Private strProperties As String


Public Property Get PWD() As String

    Set PWD = strPWD
End Property


Public Property Let PWD(val As String)

    Set strPWD = val
End Property


Public Property Get Properties() As String

    Set Properties = strProperties
End Property


Public Property Let Properties(val As String)

    Set strProperties = val
End Property


Private Sub Class_Initialize()



    strProperties = _
        "UserInterFaceOnly:=True, " & _
        "AllowFormattingCells:=True, " & _
        "AllowDeletingRows:=True, " & _
        "AllowFormattingRows:=True, " & _
        "AllowInsertingRows:=True, " & _
        "AllowSorting:=False, " & _
        "AllowFiltering:=True"

End Sub
