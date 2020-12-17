VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "clsPullCst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Declare private class constants
Private Const varHeaders As Variant = Array("PROGRAM DESCRIPTION", _
    "START DATE","END DATE","LEAD VA","LEAD CA","VA TYPE","VENDOR", _
    "BB FORMAT","COST BASIS","CA TYPE","REBATE BASIS","PRE APPROVAL", _
    "APPROP NAME","PRN/GRP","PACKET","PACKET DL","COMMENTS","","","","","")


'*******************************************************************************
'Property to return class sheet name constant (Programs)
'*******************************************************************************
Public Property Get Headers() As String

    'Return class sheet name
    Headers = varHeaders
End Property


'*******************************************************************************
'Initialize sub. Create dictionaries with column index & name keys
'*******************************************************************************
Private Sub Class_Initialize()
End Sub
