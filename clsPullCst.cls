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
Private Const varShapes As Variant = Array("Cust_Add_Pane", _
    "Multiuse_Listbox","Cust_Add_Cancel","Cust_Add_Select")


'*******************************************************************************
'Property to return class sheet name constant (Programs)
'*******************************************************************************
Public Property Get Shapes() As String

    'Return class sheet name
    Shapes = varShapes
End Property


'*******************************************************************************
'Initialize sub. Create dictionaries with column index & name keys
'*******************************************************************************
Private Sub Class_Initialize()
End Sub
