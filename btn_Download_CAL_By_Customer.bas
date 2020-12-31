Attribute VB_Name = "btn_Download_CAL_By_Customer"

'Declare private module constants TESTING3
Private Const varShp As Variant = _Array("Listbox_Pane", _"Multiuse_Listbox", _
    "Listbox_Cancel","Listbox_Select","Listbox_Account_Tgl","Listbox_All")


'*******************************************************************************
'Show all utility elements, update and resize listbox.
'*******************************************************************************
Private Sub Initialize()

    'Hide any visible shapes
    Utility.ClearShapes

    'Set default listbox view
    btnViewByAccount

    'Show utility elements
    Utility.Show(varShp)

    'Set listbox modifier buttons to this module macro
    Sheets("Control Panel").Shapes("Listbox_Account_Tgl").OnAction = _
        "btnViewByAccount"
End Sub


'*******************************************************************************
'Download CAL for account selections.
'*******************************************************************************
Sub btnSelect()

    'Declare sub variables
    Dim varCst As Variant
    Dim strCst As String
    Dim strPth As String
    Dim wb As Workbook

    'Get folder path from user selection
    strPth = Utility.SelectFolder

    'Get array of selected customer(s)
    varCst = Utility.GetSelection

    'Alert user and exit sub if no assigned customers/selected folder
    If Not IsEmpty(varCst) And strPth <> "" Then

        'Loop through all assigned customers
        For each cst In varCst

            'Create customer friendly CAL Workbook and set to variable
            Set wb = Utility.DwnCstCAL(cst)

            'Save and close workbook
            wb.Close SaveChanges:=True, Filename:= _
                strPth & cst & " CUSTOMER AGREEMENT LIST.xlsx"
        Next
    End If
End Sub


'*******************************************************************************
'Clear utility shapes from Control Panel.
'*******************************************************************************
Private Sub btnCancel()

    'Clear utility shapes from Control Panel
    Utility.ClearShapes
End Sub


'*******************************************************************************
'Update multiuse listbox with (unassigned) customer list
'*******************************************************************************
Private Sub btnViewByAccount()

    'Clear utility shapes from Control Panel
    Utility.UpdateListboxCst(True)
End Sub
