VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'*******************************************************************************
'On workbook Save/SaveAs call Push sub to insert/update all new records.
'*******************************************************************************
Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    'Update/insert all new recordset
    Push_Maintenance.Push
End Sub
